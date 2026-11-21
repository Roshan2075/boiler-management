
import React, {useEffect, useMemo, useState} from 'react';
import {createClient} from '@supabase/supabase-js';
import * as XLSX from 'xlsx-js-style';
import {
  LayoutDashboard, ClipboardList, Receipt, FileText, FolderOpen,
  LogOut, Upload, Download, Plus, RefreshCw, LockKeyhole,
  Factory, Menu, X, Trash2, Users
} from 'lucide-react';

const url = import.meta.env.VITE_SUPABASE_URL;
const key = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY || import.meta.env.VITE_SUPABASE_ANON_KEY;
const supabase = url && key ? createClient(url, key) : null;

const ADMIN_EMAIL = 'rdansena25@gmail.com';

const DEPTS = [
  'Structure', 'Pressure Part', 'Non Pressure Part',
  'Rotary Part', 'Bunker', 'LP Piping', 'Trim Piping'
];

const MONTHS = [
  'January','February','March','April','May','June',
  'July','August','September','October','November','December'
];

const blank = {
  work_date: new Date().toISOString().slice(0,10),
  department: 'Structure',
  pg: '', ma: '', du: '', description: '',
  quantity: 0, unit_weight: 0, weight: 0,
  pre_assembly: 0, placement: 0, alignment: 0,
  bolting: 0, remark: ''
};

function num(v) {
  const n = Number(String(v ?? 0).replace(/,/g,''));
  return Number.isFinite(n) ? n : 0;
}

function weight(q,u) {
  return num(q) * num(u);
}

function norm(v) {
  return String(v ?? '').toLowerCase().replace(/[^a-z0-9]/g,'');
}

function deptName(v) {
  const n = norm(v);
  return DEPTS.find(d => norm(d) === n) || '';
}

function excelDate(v) {
  if (v instanceof Date && !isNaN(v)) {
    return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}`;
  }
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }
  const s = String(v ?? '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  let m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
  if (m) {
    let y = Number(m[3]);
    if (y < 100) y += 2000;
    return `${y}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  }
  m = s.match(/^(\d{1,2})[\s-]+([A-Za-z]{3,})[\s-]+(\d{2,4})$/);
  if (m) {
    const months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
    const mo = months.findIndex(x => m[2].toLowerCase().startsWith(x)) + 1;
    let y = Number(m[3]);
    if (y < 100) y += 2000;
    if (mo) return `${y}-${String(mo).padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  }
  return '';
}

function App() {
  const [session,setSession] = useState(null);
  const [email,setEmail] = useState('');
  const [password,setPassword] = useState('');
  const [signup,setSignup] = useState(false);
  const [tab,setTab] = useState('dashboard');
  const [rows,setRows] = useState([]);
  const [bills,setBills] = useState([]);
  const [customerBills,setCustomerBills] = useState([]);
  const [docs,setDocs] = useState([]);
  const [form,setForm] = useState(blank);
  const [customer,setCustomer] = useState({
    bill_no:'', customer_name:'',
    bill_date:new Date().toISOString().slice(0,10),
    description:'', quantity:0, rate:0, gst:18
  });
  const [filterMonth,setFilterMonth] = useState(new Date().getMonth());
  const [filterYear,setFilterYear] = useState(new Date().getFullYear());
  const [filterDept,setFilterDept] = useState('All');
  const [notice,setNotice] = useState('');
  const [busy,setBusy] = useState(false);
  const [mobile,setMobile] = useState(false);
  const [stats,setStats] = useState(null);

  useEffect(() => {
    if (!supabase) return;
    supabase.auth.getSession().then(({data}) => setSession(data.session));
    const {data:{subscription}} = supabase.auth.onAuthStateChange((_e,s) => setSession(s));
    return () => subscription.unsubscribe();
  },[]);

  useEffect(() => {
    if (!session) return;
    loadAll();
    supabase.rpc('record_user_activity');
    const timer = setInterval(() => {
      supabase.rpc('record_user_activity');
    }, 60000);
    return () => clearInterval(timer);
  },[session]);

  async function loadAll() {
    if (!session) return;
    setBusy(true);
    const [d,b,c,f] = await Promise.all([
      supabase.from('dpr_entries').select('*').order('work_date',{ascending:false}),
      supabase.from('ra_bills').select('*').order('created_at',{ascending:false}),
      supabase.from('customer_bills').select('*').order('created_at',{ascending:false}),
      supabase.from('drawings').select('*').order('created_at',{ascending:false})
    ]);
    if (d.error) setNotice('DPR Load Error: '+d.error.message);
    else setRows(d.data || []);
    if (b.error) setNotice('Vendor Bill Load Error: '+b.error.message);
    else setBills(b.data || []);
    if (c.error) setNotice('Customer Bill Load Error: '+c.error.message);
    else setCustomerBills(c.data || []);
    if (f.error) setNotice('Drawing Load Error: '+f.error.message);
    else setDocs(f.data || []);
    if (session.user.email?.toLowerCase() === ADMIN_EMAIL) {
      const s = await supabase.rpc('get_admin_user_stats');
      if (!s.error) setStats(s.data);
    }
    setBusy(false);
  }

  async function login(e) {
    e.preventDefault();
    setBusy(true);
    const action = signup
      ? supabase.auth.signUp({email,password})
      : supabase.auth.signInWithPassword({email,password});
    const {error} = await action;
    setNotice(error ? error.message : signup
      ? 'Account created. Email confirmation may be required before login.'
      : 'Login successful');
    setBusy(false);
  }

  async function adminToken() {
    const p = window.prompt('Admin Password डालें:');
    if (!p) return null;
    const {data,error} = await supabase.rpc('verify_admin_password',{p_password:p});
    if (error) {
      setNotice('Admin Password Error: '+error.message);
      return null;
    }
    return data;
  }

  async function setAdminPassword() {
    const p = window.prompt('नया Admin Password डालें (कम से कम 8 अक्षर):');
    if (!p) return;
    if (p.length < 8) {
      setNotice('Password कम से कम 8 अक्षर का होना चाहिए।');
      return;
    }
    const confirmPassword = window.prompt('नया Password दोबारा डालें:');
    if (p !== confirmPassword) {
      setNotice('दोनों Password अलग हैं।');
      return;
    }
    const {error} = await supabase.rpc('set_admin_password',{p_password:p});
    setNotice(error ? error.message : 'Admin Password सफलतापूर्वक सेट हो गया।');
  }

  async function saveDpr(e) {
    e.preventDefault();
    if (!form.work_date || !form.pg || !form.ma || !form.du) {
      setNotice('Date, PG, MA और DU भरना जरूरी है।');
      return;
    }
    const token = await adminToken();
    if (!token) return;
    const payload = {
      work_date:form.work_date,
      department:form.department,
      pg:String(form.pg).trim(),
      ma:String(form.ma).trim(),
      du:String(form.du).trim(),
      description:String(form.description || ''),
      quantity:num(form.quantity),
      unit_weight:num(form.unit_weight),
      weight:weight(form.quantity,form.unit_weight),
      pre_assembly:num(form.pre_assembly),
      placement:num(form.placement),
      alignment:num(form.alignment),
      bolting:num(form.bolting),
      remark:String(form.remark || '')
    };
    const {error} = await supabase.rpc('save_dpr_entry',{
      p_token:token, p_payload:payload
    });
    setNotice(error ? error.message : 'DPR saved / updated successfully');
    if (!error) {
      setForm({...blank,work_date:form.work_date,department:form.department});
      await loadAll();
    }
  }

  async function delDpr(id) {
    if (!window.confirm('इस DPR को Delete करना है?')) return;
    const token = await adminToken();
    if (!token) return;
    const {error} = await supabase.rpc('delete_dpr_entry',{p_token:token,p_id:id});
    setNotice(error ? error.message : 'DPR deleted');
    if (!error) await loadAll();
  }

  const monthRows = useMemo(() => rows.filter(r => {
    const d = new Date(r.work_date+'T00:00:00');
    return d.getMonth() === filterMonth && d.getFullYear() === filterYear;
  }),[rows,filterMonth,filterYear]);

  const visibleRows = useMemo(() =>
    filterDept === 'All' ? monthRows : monthRows.filter(r => r.department === filterDept),
    [monthRows,filterDept]
  );

  const totals = useMemo(() => rows.reduce((a,r) => ({
    qty:a.qty+num(r.quantity),
    weight:a.weight+num(r.weight)
  }),{qty:0,weight:0}),[rows]);

  function exportRows(data,name) {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,'Records');
    XLSX.writeFile(wb,name+'.xlsx');
  }

  function exportDprWorkbook() {
    const wb = XLSX.utils.book_new();
    const headers = [
      'Date','PG','MA','DU','Description','Quantity','Unit Weight',
      'Total Weight','Pre-Assembly Qty','Pre-Assembly Weight',
      'Placement Qty','Placement Weight','Alignment Qty',
      'Alignment Weight','Bolting Qty','Bolting Weight','Remarks'
    ];
    for (const dept of DEPTS) {
      const data = rows.filter(r => r.department === dept);
      const aoa = [
        ['BOILER CONSTRUCTION MANAGEMENT'],
        [`DAILY PROGRESS REPORT - ${dept}`],
        ['Generated: '+new Date().toLocaleDateString('en-IN')],
        headers
      ];
      data.forEach(r => aoa.push([
        r.work_date,r.pg,r.ma,r.du,r.description,num(r.quantity),
        num(r.unit_weight),num(r.weight),num(r.pre_assembly),
        weight(r.pre_assembly,r.unit_weight),num(r.placement),
        weight(r.placement,r.unit_weight),num(r.alignment),
        weight(r.alignment,r.unit_weight),num(r.bolting),
        weight(r.bolting,r.unit_weight),r.remark || ''
      ]));
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = [15,12,12,12,30,14,14,16,17,19,15,18,15,18,14,17,28].map(w => ({wch:w}));
      ws['!merges'] = [
        {s:{r:0,c:0},e:{r:0,c:16}},
        {s:{r:1,c:0},e:{r:1,c:16}},
        {s:{r:2,c:0},e:{r:2,c:16}}
      ];
      XLSX.utils.book_append_sheet(wb,ws,dept);
    }
    XLSX.writeFile(wb,'Boiler_Construction_DPR.xlsx');
  }

  async function importDpr(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const token = await adminToken();
    if (!token) { e.target.value=''; return; }
    setBusy(true);
    try {
      const wb = XLSX.read(await file.arrayBuffer(),{cellDates:true});
      const mapped = [];
      let skipped = 0;
      for (const sheetName of wb.SheetNames) {
        const sheet = wb.Sheets[sheetName];
        const matrix = XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:true});
        let hi = -1, headers = [];
        for (let i=0;i<Math.min(matrix.length,25);i++) {
          const h = matrix[i].map(norm);
          if (h.includes('pg') && h.includes('ma') && h.includes('du') &&
              h.some(x => ['date','workdate'].includes(x))) {
            hi=i; headers=h; break;
          }
        }
        if (hi < 0) continue;
        const sheetDept = deptName(sheetName);
        const get = (r,aliases) => {
          const i = headers.findIndex(h => aliases.includes(h));
          return i < 0 ? '' : r[i];
        };
        for (const r of matrix.slice(hi+1)) {
          if (!r.some(v => String(v ?? '').trim())) continue;
          const date = excelDate(get(r,['date','workdate']));
          const department = sheetDept || deptName(get(r,['department','dept']));
          const pg=get(r,['pg']), ma=get(r,['ma']), du=get(r,['du']);
          if (!date || !DEPTS.includes(department) || !pg || !ma || !du) {
            skipped++; continue;
          }
          const quantity=num(get(r,['quantity','qty']));
          const unit_weight=num(get(r,['unitweight']));
          mapped.push({
            work_date:date,department,
            pg:String(pg).trim(),ma:String(ma).trim(),du:String(du).trim(),
            description:String(get(r,['description','itemdescription']) || ''),
            quantity,unit_weight,weight:quantity*unit_weight,
            pre_assembly:num(get(r,['preassemblyqty','preassembly'])),
            placement:num(get(r,['placementqty','placement'])),
            alignment:num(get(r,['alignmentqty','alignment'])),
            bolting:num(get(r,['boltingqty','bolting'])),
            remark:String(get(r,['remarks','remark','status']) || '')
          });
        }
      }
      if (!mapped.length) throw new Error('Excel में सही Date, PG, MA, DU और Department headers नहीं मिले।');
      const {data,error} = await supabase.rpc('import_dpr_entries',{
        p_token:token,p_rows:mapped
      });
      if (error) throw error;
      setNotice(`${data} DPR rows imported/updated. Skipped: ${skipped}`);
      await loadAll();
    } catch(err) {
      setNotice('Excel Import Error: '+err.message);
    } finally {
      setBusy(false);
      e.target.value='';
    }
  }

  async function createVendorBill() {
    const token = await adminToken();
    if (!token) return;
    const totals = monthRows.reduce((a,r) => ({
      quantity:a.quantity+num(r.quantity),
      weight:a.weight+num(r.weight)
    }),{quantity:0,weight:0});
    const {error} = await supabase.rpc('create_vendor_bill',{
      p_token:token,
      p_payload:{
        bill_month:MONTHS[filterMonth],
        bill_year:filterYear,
        quantity:totals.quantity,
        weight:totals.weight
      }
    });
    setNotice(error ? error.message : 'Vendor Bill saved');
    if (!error) await loadAll();
  }

  async function saveCustomerBill(e) {
    e.preventDefault();
    const token = await adminToken();
    if (!token) return;
    const q=num(customer.quantity), rate=num(customer.rate), gst=num(customer.gst);
    const amount=q*rate;
    const payload={
      bill_no:customer.bill_no,
      customer_name:customer.customer_name,
      bill_date:customer.bill_date,
      description:customer.description,
      quantity:q,rate,amount,gst,
      gst_amount:amount*gst/100,
      total_amount:amount*(1+gst/100)
    };
    const {error} = await supabase.rpc('create_customer_bill',{
      p_token:token,p_payload:payload
    });
    setNotice(error ? error.message : 'Customer Bill saved');
    if (!error) {
      setCustomer({...customer,bill_no:'',customer_name:'',description:'',quantity:0,rate:0});
      await loadAll();
    }
  }

  async function uploadDrawing(e) {
    const file=e.target.files?.[0];
    if (!file) return;
    setBusy(true);
    const path=`${session.user.id}/${Date.now()}-${file.name}`;
    const up=await supabase.storage.from('drawings').upload(path,file);
    if (up.error) {
      setNotice(up.error.message);
      setBusy(false);
      e.target.value='';
      return;
    }
    const {error}=await supabase.rpc('create_drawing_record',{
      p_payload:{
        file_name:file.name,
        storage_path:path,
        department:form.department || 'General'
      }
    });
    setNotice(error ? 'Drawing record error: '+error.message : 'Drawing uploaded successfully');
    if (error) await supabase.storage.from('drawings').remove([path]);
    else await loadAll();
    setBusy(false);
    e.target.value='';
  }

  if (!supabase) return <Setup/>;

  if (!session) return (
    <div className="login-wrap">
      <form className="login-card" onSubmit={login}>
        <div className="logo"><Factory size={34}/></div>
        <h1>Boiler Management</h1>
        <p>Construction Management System</p>
        <label>Email</label>
        <input type="email" value={email} onChange={e=>setEmail(e.target.value)} required placeholder="Email"/>
        <label>Password</label>
        <input type="password" value={password} onChange={e=>setPassword(e.target.value)} required placeholder="Password"/>
        <button className="primary full" disabled={busy}>
          <LockKeyhole size={17}/> {signup ? 'Create Account' : 'Secure Login'}
        </button>
        <button type="button" className="link" onClick={()=>setSignup(!signup)}>
          {signup ? 'पहले से Account है? Login करें' : 'नया User ID बनाएँ / Sign Up'}
        </button>
        {notice && <small>{notice}</small>}
      </form>
    </div>
  );

  const nav = [
    ['dashboard','Dashboard',LayoutDashboard],
    ['dpr','DPR Register',ClipboardList],
    ['ra','Vendor Bills',Receipt],
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
          <button key={id} className={'nav '+(tab===id?'active':'')}
            onClick={()=>{setTab(id);setMobile(false);}}>
            <Icon size={18}/>{label}
          </button>
        ))}
        <div className="side-bottom">
          <span>{session.user.email}</span>
          <button className="nav" onClick={()=>supabase.auth.signOut()}>
            <LogOut size={18}/>Logout
          </button>
        </div>
      </aside>

      <main>
        <header>
          <button className="menu" onClick={()=>setMobile(!mobile)}><Menu/></button>
          <div>
            <b>{nav.find(n=>n[0]===tab)?.[1] || 'Dashboard'}</b>
            <small>Boiler Construction Project Workspace</small>
          </div>
          <button className="icon-btn" onClick={loadAll} disabled={busy}><RefreshCw size={18}/></button>
        </header>

        <div className="content">
          {notice && <div className="notice" onClick={()=>setNotice('')}>{notice}　×</div>}

          {session.user.email?.toLowerCase()===ADMIN_EMAIL && (
            <div className="panel">
              <button className="secondary" onClick={setAdminPassword}>
                <LockKeyhole size={16}/> Set / Change Admin Password
              </button>
              {stats && <p>
                <Users size={16}/> Registered Users: <b>{stats.total_users}</b>
                {' '} | Active Users (last 5 min): <b>{stats.active_users}</b>
              </p>}
            </div>
          )}

          {tab==='dashboard' && <>
            <div className="welcome">
              <div><span className="eyebrow">PROJECT OVERVIEW</span>
                <h1>Master Dashboard</h1>
                <p>All construction records in one place.</p></div>
              <button onClick={()=>setTab('dpr')} className="primary"><Plus size={17}/> Add DPR</button>
            </div>
            <div className="stats">
              <Stat title="DPR Entries" value={rows.length}/>
              <Stat title="Total Quantity" value={totals.qty.toLocaleString()}/>
              <Stat title="Total Weight" value={totals.weight.toLocaleString()}/>
              <Stat title="Vendor Bills" value={bills.length}/>
              <Stat title="Customer Bills" value={customerBills.length}/>
              <Stat title="Drawings" value={docs.length}/>
            </div>
            <div className="panel">
              <div className="panel-head"><h2>Department Progress</h2><span>All-time quantity & weight</span></div>
              <div className="dept-grid">
                {DEPTS.map(d=>{
                  const a=rows.filter(r=>r.department===d);
                  return <div className="dept" key={d}>
                    <b>{d}</b>
                    <strong>{a.reduce((s,r)=>s+num(r.quantity),0).toLocaleString()}</strong>
                    <small>{a.length} entries · {a.reduce((s,r)=>s+num(r.weight),0).toLocaleString()} weight</small>
                  </div>;
                })}
              </div>
            </div>
            <div className="panel">
              <h2>Recent DPR Entries</h2>
              <DprTable rows={rows.slice(0,8)} onDelete={delDpr}/>
            </div>
          </>}

          {tab==='dpr' && <>
            <div className="welcome">
              <div><span className="eyebrow">DAILY PROGRESS REPORT</span><h1>DPR Register</h1></div>
              <div className="actions">
                <label className="secondary upload"><Upload size={16}/> Import Excel
                  <input type="file" accept=".xlsx,.xls,.csv" onChange={importDpr}/>
                </label>
                <button className="secondary" onClick={exportDprWorkbook}><Download size={16}/> Export Excel</button>
              </div>
            </div>
            <div className="panel">
              <h2>Add Daily Progress</h2>
              <form className="entry-grid" onSubmit={saveDpr}>
                <Field label="Date"><input type="date" value={form.work_date} onChange={e=>setForm({...form,work_date:e.target.value})} required/></Field>
                <Field label="Department"><select value={form.department} onChange={e=>setForm({...form,department:e.target.value})}>{DEPTS.map(d=><option key={d}>{d}</option>)}</select></Field>
                {['pg','ma','du','description','quantity','unit_weight','pre_assembly','placement','alignment','bolting','remark'].map(k=>
                  <Field label={k.replaceAll('_',' ').toUpperCase()} key={k}>
                    <input type={['quantity','unit_weight','pre_assembly','placement','alignment','bolting'].includes(k)?'number':'text'}
                      step="any" value={form[k]} onChange={e=>setForm({...form,[k]:e.target.value})}/>
                  </Field>
                )}
                <Field label="TOTAL WEIGHT (AUTO)"><input value={weight(form.quantity,form.unit_weight)} readOnly/></Field>
                <div className="form-end"><button className="primary" disabled={busy}><Plus size={16}/> Save / Update DPR</button></div>
              </form>
            </div>
            <div className="panel">
              <div className="panel-head">
                <h2>DPR Register</h2>
                <div className="filters">
                  <select value={filterDept} onChange={e=>setFilterDept(e.target.value)}>
                    <option>All</option>{DEPTS.map(d=><option key={d}>{d}</option>)}
                  </select>
                  <select value={filterMonth} onChange={e=>setFilterMonth(+e.target.value)}>
                    {MONTHS.map((m,i)=><option key={m} value={i}>{m}</option>)}
                  </select>
                  <input type="number" value={filterYear} onChange={e=>setFilterYear(+e.target.value)}/>
                </div>
              </div>
              <div className="month-summary">
                <b>{MONTHS[filterMonth]} {filterYear}</b>
                <span>{visibleRows.length} entries</span>
                <span>Qty: {visibleRows.reduce((s,r)=>s+num(r.quantity),0)}</span>
                <span>Weight: {visibleRows.reduce((s,r)=>s+num(r.weight),0)}</span>
              </div>
              <DprTable rows={visibleRows} onDelete={delDpr}/>
            </div>
          </>}

          {tab==='ra' && <>
            <div className="welcome">
              <div><span className="eyebrow">VENDOR BILLING</span><h1>Vendor Bill Management</h1>
                <p>Generate monthly quantity and weight summaries from DPR.</p></div>
              <button className="primary" onClick={createVendorBill}><Plus size={17}/> Create Monthly Vendor Bill</button>
            </div>
            <div className="panel">
              <div className="panel-head"><h2>Selected DPR Period</h2>
                <div className="filters">
                  <select value={filterMonth} onChange={e=>setFilterMonth(+e.target.value)}>{MONTHS.map((m,i)=><option key={m} value={i}>{m}</option>)}</select>
                  <input type="number" value={filterYear} onChange={e=>setFilterYear(+e.target.value)}/>
                </div>
              </div>
              <p>Linked DPR totals: <b>{monthRows.reduce((s,r)=>s+num(r.quantity),0)} quantity</b> · <b>{monthRows.reduce((s,r)=>s+num(r.weight),0)} weight</b></p>
              <button className="secondary" onClick={()=>exportRows(monthRows,`DPR_${MONTHS[filterMonth]}_${filterYear}_For_Vendor`)}><Download size={16}/> Export linked DPR</button>
            </div>
            <div className="panel">
              <h2>Saved Vendor Bills</h2>
              <DataTable rows={bills} columns={['bill_month','bill_year','quantity','weight','created_at']}/>
              <button className="secondary" onClick={()=>exportRows(bills,'Vendor_Bills')}><Download size={16}/> Download Vendor Bills</button>
            </div>
          </>}

          {tab==='customer' && <>
            <div className="welcome">
              <div><span className="eyebrow">CUSTOMER BILLING</span><h1>Customer Bills</h1></div>
              <button className="secondary" onClick={()=>exportRows(customerBills,'Customer_Bills')}><Download size={16}/> Export Excel</button>
            </div>
            <div className="panel">
              <h2>Create Customer Bill</h2>
              <form className="entry-grid" onSubmit={saveCustomerBill}>
                {Object.keys(customer).map(k=><Field label={k.replaceAll('_',' ').toUpperCase()} key={k}>
                  <input type={['quantity','rate','gst'].includes(k)?'number':k==='bill_date'?'date':'text'}
                    step="any" value={customer[k]} required
                    onChange={e=>setCustomer({...customer,[k]:e.target.value})}/>
                </Field>)}
                <div className="form-end">
                  <b>Estimated Total: ₹{(num(customer.quantity)*num(customer.rate)*(1+num(customer.gst)/100)).toFixed(2)}</b>
                  <button className="primary"><Plus size={16}/> Save Customer Bill</button>
                </div>
              </form>
            </div>
            <div className="panel"><h2>Customer Bill Register</h2>
              <DataTable rows={customerBills} columns={['bill_no','customer_name','bill_date','description','quantity','rate','amount','gst_amount','total_amount']}/>
            </div>
          </>}

          {tab==='drawings' && <>
            <div className="welcome">
              <div><span className="eyebrow">DOCUMENT LIBRARY</span><h1>Drawing Folder</h1>
                <p>Store and organize drawings by department.</p></div>
              <label className="primary upload"><Upload size={16}/> Upload Drawing
                <input type="file" accept=".pdf,.dwg,.dxf,.png,.jpg,.jpeg,.xlsx" onChange={uploadDrawing}/>
              </label>
            </div>
            <div className="panel">
              <Field label="Drawing Department">
                <select value={form.department} onChange={e=>setForm({...form,department:e.target.value})}>
                  {DEPTS.map(d=><option key={d}>{d}</option>)}
                </select>
              </Field>
              <div className="dept-grid">
                {DEPTS.map(d=><div className="dept" key={d}><b>{d}</b><small>{docs.filter(x=>x.department===d).length} files</small></div>)}
              </div>
            </div>
            <div className="panel">
              <h2>Saved Drawings</h2>
              {docs.map(f=><div className="file-row" key={f.id}>
                <FolderOpen/>
                <div><b>{f.file_name}</b><small>{f.department}</small></div>
                <button className="link" onClick={async()=>{
                  const {data,error}=await supabase.storage.from('drawings').createSignedUrl(f.storage_path,120);
                  if(error) setNotice(error.message);
                  else window.open(data.signedUrl,'_blank');
                }}>Open</button>
              </div>)}
              {!docs.length && <div className="empty">No drawings uploaded yet.</div>}
            </div>
          </>}
        </div>
      </main>
    </div>
  );
}

function Setup() {
  return <div className="login-wrap"><div className="login-card">
    <Factory size={38}/><h1>Boiler Management</h1>
    <p>Supabase connection required</p>
    <div className="setup-error">Project URL या Publishable Key सेट नहीं है।</div>
  </div></div>;
}

function Stat({title,value}) {
  return <div className="stat"><small>{title}</small><strong>{value}</strong></div>;
}

function Field({label,children}) {
  return <label className="field"><span>{label}</span>{children}</label>;
}

function DprTable({rows,onDelete}) {
  return <div className="table-wrap"><table>
    <thead><tr>{['Date','Dept.','PG','MA','DU','Description','Qty','Unit Weight','Total Weight','Pre-Assembly','Placement','Alignment','Bolting','Remarks',''].map(x=><th key={x}>{x}</th>)}</tr></thead>
    <tbody>{rows.map(r=><tr key={r.id}>
      <td>{r.work_date}</td><td>{r.department}</td><td>{r.pg}</td><td>{r.ma}</td><td>{r.du}</td>
      <td>{r.description}</td><td>{r.quantity}</td><td>{r.unit_weight ?? 0}</td>
      <td>{weight(r.quantity,r.unit_weight)}</td><td>{r.pre_assembly}</td>
      <td>{r.placement}</td><td>{r.alignment}</td><td>{r.bolting}</td><td>{r.remark}</td>
      <td><button className="delete" onClick={()=>onDelete(r.id)}><Trash2 size={15}/></button></td>
    </tr>)}</tbody>
  </table>{!rows.length && <div className="empty">No records found.</div>}</div>;
}

function DataTable({rows,columns}) {
  return <div className="table-wrap"><table>
    <thead><tr>{columns.map(c=><th key={c}>{c.replaceAll('_',' ')}</th>)}</tr></thead>
    <tbody>{rows.map(r=><tr key={r.id || r.created_at}>
      {columns.map(c=><td key={c}>{r[c]}</td>)}
    </tr>)}</tbody>
  </table>{!rows.length && <div className="empty">No records yet.</div>}</div>;
}

export default App;

