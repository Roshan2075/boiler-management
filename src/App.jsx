import React,{useEffect,useState} from 'react';
import {createClient} from '@supabase/supabase-js';
import * as XLSX from 'xlsx-js-style';
import {
 LayoutDashboard,ClipboardList,Receipt,FileText,FolderOpen,
 LogOut,Upload,Download,Plus,RefreshCw,LockKeyhole,
 Factory,Menu,X,Trash2,Eye,Printer,Building2
} from 'lucide-react';
import './style.css';

const supabase=createClient(
 import.meta.env.VITE_SUPABASE_URL||'',
 import.meta.env.VITE_SUPABASE_ANON_KEY||import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY||''
);
const COMPANY='P.E. ERECTORS PVT. LTD.';
const DEPTS=['Structure','Pressure Part','Non Pressure Part','Rotary Part','Bunker','LP Piping','Trim Piping'];
const money=n=>'₹ '+(Number(n)||0).toLocaleString('en-IN',{minimumFractionDigits:2,maximumFractionDigits:2});
const today=()=>new Date().toISOString().slice(0,10);
const blankDpr={work_date:today(),department:'Structure',pg:'',ma:'',du:'',description:'',quantity:0,weight:0,pre_assembly:0,placement:0,alignment:0,bolting:0,remark:''};
const blankBill={agency_name:'',bill_no:'',bill_date:today(),period_from:'',period_to:'',previous_amount:0,current_amount:0,cumulative_amount:0,gst_amount:0,gross_bill:0,deductions:0,net_payable:0,abstract_json:[]};

function App(){
 const [session,setSession]=useState(null);
 const [email,setEmail]=useState('');
 const [password,setPassword]=useState('');
 const [tab,setTab]=useState('dashboard');
 const [mobile,setMobile]=useState(false);
 const [busy,setBusy]=useState(false);
 const [notice,setNotice]=useState('');
 const [dpr,setDpr]=useState([]);
 const [vendor,setVendor]=useState([]);
 const [customer,setCustomer]=useState([]);
 const [drawings,setDrawings]=useState([]);
 const [dform,setDform]=useState(blankDpr);
 const [vform,setVform]=useState(blankBill);
 const [cform,setCform]=useState({...blankBill,agency_name:'BHEL PSER LARA SITE'});
 const [view,setView]=useState(null);
 const [uploadDept,setUploadDept]=useState('Structure');

 useEffect(()=>{
  supabase.auth.getSession().then(({data})=>setSession(data.session));
  const {data:{subscription}}=supabase.auth.onAuthStateChange((_e,s)=>setSession(s));
  return()=>subscription.unsubscribe();
 },[]);

 useEffect(()=>{if(session)loadAll()},[session]);

 async function loadAll(){
  setBusy(true);
  const specs=[
   ['dpr_entries',setDpr,'work_date'],
   ['vendor_bills',setVendor,'created_at'],
   ['customer_bill_abstracts',setCustomer,'created_at'],
   ['drawings',setDrawings,'created_at']
  ];
  for(const [table,setter,sort] of specs){
   const {data,error}=await supabase.from(table).select('*').order(sort,{ascending:false});
   if(!error)setter(data||[]);
   else setNotice(table+': '+error.message);
  }
  setBusy(false);
 }

 async function login(e){
  e.preventDefault();
  setBusy(true);
  const {error}=await supabase.auth.signInWithPassword({email,password});
  setNotice(error?.message||'Login successful');
  setBusy(false);
 }

 async function saveDpr(e){
  e.preventDefault();
  const payload={
   ...dform,
   quantity:+dform.quantity||0,
   weight:+dform.weight||0,
   pre_assembly:+dform.pre_assembly||0,
   placement:+dform.placement||0,
   alignment:+dform.alignment||0,
   bolting:+dform.bolting||0,
   user_id:session.user.id
  };
  const {error}=await supabase.from('dpr_entries').upsert(payload,{
   onConflict:'work_date,department,pg,ma,du'
  });
  setNotice(error?.message||'DPR saved / updated');
  if(!error){
   setDform({...blankDpr,work_date:dform.work_date,department:dform.department});
   loadAll();
  }
 }

 async function deleteRow(table,id){
  if(!confirm('क्या आप यह रिकॉर्ड हटाना चाहते हैं?'))return;
  const {error}=await supabase.from(table).delete().eq('id',id);
  setNotice(error?.message||'Record deleted');
  if(!error)loadAll();
 }

 function billPayload(f){
  return {
   ...f,
   previous_amount:+f.previous_amount||0,
   current_amount:+f.current_amount||0,
   cumulative_amount:+f.cumulative_amount||0,
   gst_amount:+f.gst_amount||0,
   gross_bill:+f.gross_bill||0,
   deductions:+f.deductions||0,
   net_payable:+f.net_payable||0,
   user_id:session.user.id
  };
 }

 async function saveBill(kind,e){
  e.preventDefault();
  const table=kind==='vendor'?'vendor_bills':'customer_bill_abstracts';
  const form=kind==='vendor'?vform:cform;
  const {error}=await supabase.from(table).insert(billPayload(form));
  setNotice(error?.message||'Bill saved successfully');
  if(!error){
   if(kind==='vendor')setVform(blankBill);
   else setCform({...blankBill,agency_name:'BHEL PSER LARA SITE'});
   loadAll();
  }
 }

 function addItem(kind){
  const setter=kind==='vendor'?setVform:setCform;
  const form=kind==='vendor'?vform:cform;
  setter({...form,abstract_json:[...(form.abstract_json||[]),{description:'',qty:0,unit:'',rate:0}]});
 }

 function updateItem(kind,i,key,val){
  const setter=kind==='vendor'?setVform:setCform;
  const form=kind==='vendor'?vform:cform;
  const items=[...(form.abstract_json||[])];
  items[i]={...items[i],[key]:val};
  setter({...form,abstract_json:items});
 }

 async function uploadDrawing(e){
  const file=e.target.files?.[0];
  if(!file)return;
  const allowed=['pdf','dwg','dxf','png','jpg','jpeg','xlsx','xls'];
  const ext=file.name.split('.').pop().toLowerCase();
  if(!allowed.includes(ext)){
   setNotice('यह file type allowed नहीं है');
   e.target.value='';
   return;
  }
  if(file.size>50*1024*1024){
   setNotice('File 50 MB से बड़ी है');
   e.target.value='';
   return;
  }
  setBusy(true);
  const path=`${session.user.id}/${Date.now()}-${file.name.replace(/[^a-zA-Z0-9._-]/g,'_')}`;
  const {error:upErr}=await supabase.storage.from('drawings').upload(path,file,{
   upsert:false,
   contentType:file.type||'application/octet-stream'
  });
  if(upErr){
   setNotice('Upload failed: '+upErr.message);
   setBusy(false);
   return;
  }
  const {error:dbErr}=await supabase.from('drawings').insert({
   file_name:file.name,
   storage_path:path,
   department:uploadDept,
   user_id:session.user.id
  });
  if(dbErr){
   await supabase.storage.from('drawings').remove([path]);
   setNotice('Record save failed: '+dbErr.message);
  }else{
   setNotice('Drawing uploaded successfully');
   loadAll();
  }
  setBusy(false);
  e.target.value='';
 }

 async function openDrawing(f){
  const {data,error}=await supabase.storage.from('drawings').createSignedUrl(f.storage_path,3600);
  if(error)setNotice(error.message);
  else window.open(data.signedUrl,'_blank','noopener,noreferrer');
 }

 function excelBill(b,kind){
  const aoa=[
   [COMPANY],['BILL ABSTRACT'],
   ['Bill Type',kind==='vendor'?'VENDOR BILL':'CUSTOMER BILL'],
   ['Agency / Customer',b.agency_name||''],
   ['Bill No.',b.bill_no||'','Bill Date',b.bill_date||''],
   ['Billing Period',`${b.period_from||''} to ${b.period_to||''}`],
   [],
   ['S.No.','Description','Quantity','Unit','Rate','Amount'],
   ...(b.abstract_json||[]).map((r,i)=>[
    i+1,r.description,r.qty,r.unit,r.rate,(+r.qty||0)*(+r.rate||0)
   ]),
   [],
   ['Previous Amount',b.previous_amount],
   ['Current Amount',b.current_amount],
   ['Cumulative Amount',b.cumulative_amount],
   ['GST',b.gst_amount],
   ['Gross Bill',b.gross_bill],
   ['Deductions',b.deductions],
   ['NET PAYABLE',b.net_payable]
  ];
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges']=[
   {s:{r:0,c:0},e:{r:0,c:5}},
   {s:{r:1,c:0},e:{r:1,c:5}}
  ];
  ws['!cols']=[
   {wch:12},{wch:42},{wch:14},{wch:12},{wch:16},{wch:18}
  ];
  const range=XLSX.utils.decode_range(ws['!ref']);
  for(let R=0;R<=range.e.r;R++){
   for(let C=0;C<=range.e.c;C++){
    const cell=ws[XLSX.utils.encode_cell({r:R,c:C})];
    if(!cell)continue;
    cell.s={
     font:{
      name:'Arial',
      sz:R<2?14:10,
      bold:R<2||R===7||R>=aoa.length-7,
      color:{rgb:R===aoa.length-1?'FFFFFF':'172B4D'}
     },
     alignment:{vertical:'center',horizontal:C>=2?'right':'left',wrapText:true},
     border:{
      top:{style:'thin',color:{rgb:'CBD5E1'}},
      bottom:{style:'thin',color:{rgb:'CBD5E1'}},
      left:{style:'thin',color:{rgb:'CBD5E1'}},
      right:{style:'thin',color:{rgb:'CBD5E1'}}
     },
     fill:{fgColor:{rgb:R===aoa.length-1?'173B67':R===7?'DCE8F5':R<2?'EAF1F8':'FFFFFF'}}
    };
   }
  }
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Bill Abstract');
  XLSX.writeFile(wb,`${kind}_${(b.bill_no||'Bill').replace(/[^a-z0-9_-]/gi,'_')}.xlsx`);
 }

 function printBill(b,kind){
  const w=window.open('','_blank');
  if(!w)return;
  const items=(b.abstract_json||[]).map((r,i)=>
   `<tr><td>${i+1}</td><td>${esc(r.description)}</td><td>${r.qty||0}</td><td>${esc(r.unit)}</td><td>${money(r.rate)}</td><td>${money((+r.qty||0)*(+r.rate||0))}</td></tr>`
  ).join('');
  w.document.write(`
   <html><head><title>${esc(b.bill_no)}</title>
   <style>
   body{font:12px Arial;padding:25px;color:#172b4d}
   h1,h2{text-align:center;margin:5px}
   table{border-collapse:collapse;width:100%;margin-top:18px}
   td,th{border:1px solid #777;padding:7px}
   th{background:#dce8f5}
   .tot{width:42%;margin-left:auto}
   .tot td:last-child{text-align:right;font-weight:bold}
   .head{display:flex;justify-content:space-between;margin:25px 0}
   </style></head><body>
   <h1>${COMPANY}</h1>
   <h2>${kind==='vendor'?'VENDOR BILL ABSTRACT':'CUSTOMER BILL ABSTRACT'}</h2>
   <div class="head">
    <span>Agency: ${esc(b.agency_name)}</span>
    <span>Bill No: ${esc(b.bill_no)}</span>
    <span>Date: ${esc(b.bill_date)}</span>
   </div>
   <p>Billing Period: ${esc(b.period_from)} to ${esc(b.period_to)}</p>
   <table><tr><th>S.No.</th><th>Description</th><th>Qty</th><th>Unit</th><th>Rate</th><th>Amount</th></tr>${items}</table>
   <table class="tot">${[
    ['Previous Amount',b.previous_amount],
    ['Current Amount',b.current_amount],
    ['Cumulative Amount',b.cumulative_amount],
    ['GST',b.gst_amount],
    ['Gross Bill',b.gross_bill],
    ['Deductions',b.deductions],
    ['NET PAYABLE',b.net_payable]
   ].map(([k,v])=>`<tr><td>${k}</td><td>${money(v)}</td></tr>`).join('')}</table>
   <script>window.onload=()=>window.print()<\/script>
   </body></html>
  `);
  w.document.close();
 }

 const totalDpr=dpr.reduce((a,r)=>a+(+r.quantity||0),0);
 const vendorTotal=vendor.reduce((a,r)=>a+(+r.net_payable||0),0);
 const customerTotal=customer.reduce((a,r)=>a+(+r.net_payable||0),0);

 const nav=[
  ['dashboard','Dashboard',LayoutDashboard],
  ['dpr','DPR Register',ClipboardList],
  ['vendor','Vendor Bills',Receipt],
  ['customer','Customer Bills',FileText],
  ['drawings','Drawings',FolderOpen]
 ];

 if(!import.meta.env.VITE_SUPABASE_URL||
 !(import.meta.env.VITE_SUPABASE_ANON_KEY||import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY)){
  return <div className="login-wrap"><div className="login-card">
   <Factory size={38}/><h1>{COMPANY}</h1>
   <p>Supabase configuration missing</p>
   <p>.env में VITE_SUPABASE_URL और VITE_SUPABASE_ANON_KEY सेट करें।</p>
  </div></div>;
 }

 if(!session){
  return <div className="login-wrap">
   <form className="login-card" onSubmit={login}>
    <div className="logo"><Building2 size={34}/></div>
    <h1>{COMPANY}</h1>
    <p>Boiler Construction Management Portal</p>
    <label>Email</label>
    <input type="email" value={email} onChange={e=>setEmail(e.target.value)} required/>
    <label>Password</label>
    <input type="password" value={password} onChange={e=>setPassword(e.target.value)} required/>
    <button className="primary full" disabled={busy}>
     <LockKeyhole size={17}/> {busy?'Please wait...':'Login Securely'}
    </button>
    {notice&&<small>{notice}</small>}
   </form>
  </div>;
 }

 return <div className="app">
  <aside className={mobile?'open':''}>
   <div className="brand">
    <Factory/>
    <span>{COMPANY}<small>BOILER MANAGEMENT</small></span>
    <button className="mobile-close" onClick={()=>setMobile(false)}><X/></button>
   </div>
   <div className="nav-label">MAIN MENU</div>
   {nav.map(([id,label,Icon])=>
    <button key={id} className={'nav '+(tab===id?'active':'')} onClick={()=>{
     setTab(id);setMobile(false);
    }}>
     <Icon size={18}/>{label}
    </button>
   )}
   <div className="side-bottom">
    <span>Secure workspace</span>
    <button className="nav" onClick={()=>supabase.auth.signOut()}>
     <LogOut size={18}/>Logout
    </button>
   </div>
  </aside>

  <main>
   <header>
    <button className="menu" onClick={()=>setMobile(!mobile)}><Menu/></button>
    <div><b>{COMPANY}</b><small>Boiler Construction Project Workspace</small></div>
    <button className="icon-btn" onClick={loadAll} disabled={busy}><RefreshCw size={18}/></button>
   </header>

   <div className="content">
    {notice&&<div className="notice" onClick={()=>setNotice('')}>{notice}　×</div>}

    {tab==='dashboard'&&<>
     <div className="welcome">
      <div><span className="eyebrow">COMPANY WORKSPACE</span>
       <h1>{COMPANY}</h1>
       <p>Project billing, daily progress and drawing records — all in one place.</p>
      </div>
      <button className="primary" onClick={()=>setTab('vendor')}><Plus size={17}/> New Bill</button>
     </div>
     <div className="stats">
      <Stat title="DPR Entries" value={dpr.length}/>
      <Stat title="Total DPR Quantity" value={totalDpr.toLocaleString('en-IN')}/>
      <Stat title="Vendor Bill Value" value={money(vendorTotal)}/>
      <Stat title="Customer Bill Value" value={money(customerTotal)}/>
     </div>
     <div className="panel">
      <div className="panel-head"><h2>Quick Actions</h2></div>
      <div className="quick-grid">
       <button onClick={()=>setTab('dpr')}><ClipboardList/> Add DPR Entry</button>
       <button onClick={()=>setTab('vendor')}><Receipt/> Create Vendor Bill</button>
       <button onClick={()=>setTab('customer')}><FileText/> Create Customer Bill</button>
       <button onClick={()=>setTab('drawings')}><Upload/> Upload Drawing</button>
      </div>
     </div>
     <div className="panel">
      <h2>Recent DPR Entries</h2>
      <DprTable rows={dpr.slice(0,8)} onDelete={id=>deleteRow('dpr_entries',id)}/>
     </div>
    </>}

    {tab==='dpr'&&<>
     <PageTitle title="DPR Register" desc="Department-wise daily progress. Same date + department + PG + MA + DU updates existing record."/>
     <div className="panel">
      <h2>Add / Update Daily Progress</h2>
      <form className="entry-grid" onSubmit={saveDpr}>
       <Field label="Date"><input type="date" value={dform.work_date} onChange={e=>setDform({...dform,work_date:e.target.value})} required/></Field>
       <Field label="Department"><select value={dform.department} onChange={e=>setDform({...dform,department:e.target.value})}>{DEPTS.map(x=><option key={x}>{x}</option>)}</select></Field>
       {['pg','ma','du','description','quantity','weight','pre_assembly','placement','alignment','bolting','remark'].map(k=>
        <Field key={k} label={k.replaceAll('_',' ').toUpperCase()}>
         <input type={['quantity','weight','pre_assembly','placement','alignment','bolting'].includes(k)?'number':'text'} step="any" value={dform[k]} onChange={e=>setDform({...dform,[k]:e.target.value})}/>
        </Field>
       )}
       <div className="form-end"><button className="primary"><Plus size={16}/> Save / Update DPR</button></div>
      </form>
     </div>
     <div className="panel">
      <div className="panel-head"><h2>All DPR Records</h2><button className="secondary" onClick={()=>exportDpr(dpr)}><Download size={16}/> Excel</button></div>
      <DprTable rows={dpr} onDelete={id=>deleteRow('dpr_entries',id)}/>
     </div>
    </>}

    {tab==='vendor'&&<BillPage kind="vendor" form={vform} setForm={setVform} save={e=>saveBill('vendor',e)} addItem={()=>addItem('vendor')} updateItem={(i,k,v)=>updateItem('vendor',i,k,v)} rows={vendor} view={b=>setView({bill:b,kind:'vendor'})} excel={b=>excelBill(b,'vendor')} print={b=>printBill(b,'vendor')} remove={id=>deleteRow('vendor_bills',id)}/>}

    {tab==='customer'&&<BillPage kind="customer" form={cform} setForm={setCform} save={e=>saveBill('customer',e)} addItem={()=>addItem('customer')} updateItem={(i,k,v)=>updateItem('customer',i,k,v)} rows={customer} view={b=>setView({bill:b,kind:'customer'})} excel={b=>excelBill(b,'customer')} print={b=>printBill(b,'customer')} remove={id=>deleteRow('customer_bill_abstracts',id)}/>}

    {tab==='drawings'&&<>
     <PageTitle title="Drawing Library" desc="Upload PDF, DWG, DXF, images and Excel drawings securely."/>
     <div className="panel">
      <div className="entry-grid">
       <Field label="Department"><select value={uploadDept} onChange={e=>setUploadDept(e.target.value)}>{DEPTS.map(x=><option key={x}>{x}</option>)}</select></Field>
       <Field label="Select Drawing"><input type="file" accept=".pdf,.dwg,.dxf,.png,.jpg,.jpeg,.xlsx,.xls" onChange={uploadDrawing}/></Field>
       <span className="hint">Maximum 50 MB. Allowed: PDF, DWG, DXF, JPG, PNG, XLSX.</span>
      </div>
     </div>
     <div className="panel">
      <h2>Uploaded Files ({drawings.length})</h2>
      {drawings.map(f=><div className="file-row" key={f.id}>
       <FolderOpen/>
       <div><b>{f.file_name}</b><small>{f.department}</small></div>
       <button className="secondary" onClick={()=>openDrawing(f)}><Eye size={15}/> See / Download</button>
       <button className="delete" onClick={()=>deleteRow('drawings',f.id)}><Trash2 size={15}/></button>
      </div>)}
     </div>
    </>}
   </div>
  </main>

  {view&&<BillPreview bill={view.bill} kind={view.kind} close={()=>setView(null)} excel={()=>excelBill(view.bill,view.kind)} print={()=>printBill(view.bill,view.kind)}/>}
 </div>;
}

function BillPage({kind,form,setForm,save,addItem,updateItem,rows,view,excel,print,remove}){
 const vendor=kind==='vendor';
 return <>
  <PageTitle title={vendor?'Vendor Bills':'Customer Bills'} desc="Enter bill values manually. Preview and Excel export use the same abstract data."/>
  <div className="panel">
   <h2>{vendor?'Create Vendor Bill':'Create Customer Bill'}</h2>
   <form onSubmit={save}>
    <div className="entry-grid">
     <Field label="Agency / Customer Name"><input value={form.agency_name} onChange={e=>setForm({...form,agency_name:e.target.value})} required/></Field>
     <Field label="Bill Number"><input value={form.bill_no} onChange={e=>setForm({...form,bill_no:e.target.value})} required/></Field>
     <Field label="Bill Date"><input type="date" value={form.bill_date} onChange={e=>setForm({...form,bill_date:e.target.value})} required/></Field>
     <Field label="Period From"><input type="date" value={form.period_from} onChange={e=>setForm({...form,period_from:e.target.value})}/></Field>
     <Field label="Period To"><input type="date" value={form.period_to} onChange={e=>setForm({...form,period_to:e.target.value})}/></Field>
    </div>

    <div className="panel-head"><h3>Bill Abstract / Items</h3><button type="button" className="secondary" onClick={addItem}><Plus size={15}/> Add Item</button></div>
    <div className="table-wrap"><table>
     <thead><tr><th>Description</th><th>Qty</th><th>Unit</th><th>Rate</th><th>Amount</th><th></th></tr></thead>
     <tbody>{(form.abstract_json||[]).map((r,i)=><tr key={i}>
      <td><input value={r.description} onChange={e=>updateItem(i,'description',e.target.value)}/></td>
      <td><input type="number" step="any" value={r.qty} onChange={e=>updateItem(i,'qty',e.target.value)}/></td>
      <td><input value={r.unit} onChange={e=>updateItem(i,'unit',e.target.value)}/></td>
      <td><input type="number" step="any" value={r.rate} onChange={e=>updateItem(i,'rate',e.target.value)}/></td>
      <td>{money((+r.qty||0)*(+r.rate||0))}</td>
      <td><button type="button" className="delete" onClick={()=>setForm({...form,abstract_json:form.abstract_json.filter((_,n)=>n!==i)})}><Trash2 size={14}/></button></td>
     </tr>)}</tbody>
    </table></div>

    <div className="entry-grid totals-grid">
     {['previous_amount','current_amount','cumulative_amount','gst_amount','gross_bill','deductions','net_payable'].map(k=>
      <Field label={k.replaceAll('_',' ').toUpperCase()} key={k}>
       <input type="number" step="0.01" value={form[k]} onChange={e=>setForm({...form,[k]:e.target.value})}/>
      </Field>
     )}
    </div>
    <div className="form-end"><button className="primary"><Plus size={16}/> Save {vendor?'Vendor':'Customer'} Bill</button></div>
   </form>
  </div>

  <div className="panel">
   <div className="panel-head"><h2>Saved {vendor?'Vendor':'Customer'} Bills</h2><span>{rows.length} records</span></div>
   <div className="table-wrap"><table>
    <thead><tr><th>Bill No.</th><th>Agency / Customer</th><th>Date</th><th>Period</th><th>Net Payable</th><th>Actions</th></tr></thead>
    <tbody>{rows.map(b=><tr key={b.id}>
     <td>{b.bill_no}</td><td>{b.agency_name}</td><td>{b.bill_date}</td>
     <td>{b.period_from} to {b.period_to}</td><td>{money(b.net_payable)}</td>
     <td className="actions">
      <button className="secondary" onClick={()=>view(b)}><Eye size={14}/> See</button>
      <button className="secondary" onClick={()=>excel(b)}><Download size={14}/> Excel</button>
      <button className="delete" onClick={()=>remove(b.id)}><Trash2 size={14}/></button>
     </td>
    </tr>)}</tbody>
   </table>
   {!rows.length&&<div className="empty">अभी कोई बिल सेव नहीं है।</div>}
   </div>
  </div>
 </>;
}

function BillPreview({bill,kind,close,excel,print}){
 return <div className="modal-back" onClick={close}>
  <div className="bill-modal" onClick={e=>e.stopPropagation()}>
   <div className="modal-actions">
    <button className="secondary" onClick={excel}><Download size={15}/> Download Excel</button>
    <button className="secondary" onClick={print}><Printer size={15}/> Print</button>
    <button className="delete" onClick={close}><X/></button>
   </div>
   <div className="bill-paper">
    <h1>{COMPANY}</h1>
    <h2>{kind==='vendor'?'VENDOR BILL ABSTRACT':'CUSTOMER BILL ABSTRACT'}</h2>
    <div className="bill-meta">
     <span><b>Agency:</b> {bill.agency_name}</span>
     <span><b>Bill No:</b> {bill.bill_no}</span>
     <span><b>Date:</b> {bill.bill_date}</span>
     <span><b>Period:</b> {bill.period_from} to {bill.period_to}</span>
    </div>
    <table><thead><tr><th>S.No.</th><th>Description</th><th>Qty</th><th>Unit</th><th>Rate</th><th>Amount</th></tr></thead>
     <tbody>{(bill.abstract_json||[]).map((r,i)=><tr key={i}>
      <td>{i+1}</td><td>{r.description}</td><td>{r.qty}</td><td>{r.unit}</td><td>{money(r.rate)}</td><td>{money((+r.qty||0)*(+r.rate||0))}</td>
     </tr>)}</tbody>
    </table>
    <table className="bill-totals"><tbody>
     {[
      ['Previous Amount',bill.previous_amount],
      ['Current Amount',bill.current_amount],
      ['Cumulative Amount',bill.cumulative_amount],
      ['GST',bill.gst_amount],
      ['Gross Bill',bill.gross_bill],
      ['Deductions',bill.deductions],
      ['NET PAYABLE',bill.net_payable]
     ].map(([k,v])=><tr key={k}><th>{k}</th><td>{money(v)}</td></tr>)}
    </tbody></table>
   </div>
  </div>
 </div>;
}

function PageTitle({title,desc}){
 return <div className="welcome"><div><span className="eyebrow">PROJECT MANAGEMENT</span><h1>{title}</h1><p>{desc}</p></div></div>;
}
function Stat({title,value}){
 return <div className="stat"><small>{title}</small><strong>{value}</strong></div>;
}
function Field({label,children}){
 return <label className="field"><span>{label}</span>{children}</label>;
}
function DprTable({rows,onDelete}){
 return <div className="table-wrap"><table>
  <thead><tr>{['Date','Department','PG','MA','DU','Description','Qty','Weight','Pre-Assembly','Placement','Alignment','Bolting','Remark',''].map(x=><th key={x}>{x}</th>)}</tr></thead>
  <tbody>{rows.map(r=><tr key={r.id}>
   <td>{r.work_date}</td><td>{r.department}</td><td>{r.pg}</td><td>{r.ma}</td><td>{r.du}</td>
   <td>{r.description}</td><td>{r.quantity}</td><td>{r.weight}</td><td>{r.pre_assembly}</td>
   <td>{r.placement}</td><td>{r.alignment}</td><td>{r.bolting}</td><td>{r.remark}</td>
   <td><button className="delete" onClick={()=>onDelete(r.id)}><Trash2 size={14}/></button></td>
  </tr>)}</tbody>
 </table>
 {!rows.length&&<div className="empty">No DPR records found.</div>}
 </div>;
}

function exportDpr(rows){
 const wb=XLSX.utils.book_new();
 for(const dept of DEPTS){
  const data=rows.filter(r=>r.department===dept);
  const ws=XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb,ws,dept.slice(0,31));
 }
 XLSX.writeFile(wb,'DPR_Seven_Departments.xlsx');
}
function esc(s){
 return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

export default App;

