import React, { useEffect, useMemo, useState } from 'react';
import { createClient } from '@supabase/supabase-js';
import * as XLSX from 'xlsx-js-style';
import {
  LayoutDashboard, ClipboardList, Receipt, FileText, FolderOpen,
  LogOut, Upload, Download, Plus, RefreshCw, LockKeyhole,
  Factory, Menu, X, Trash2, Users, Eye, Printer, IndianRupee
} from 'lucide-react';

const url = import.meta.env.VITE_SUPABASE_URL;
const key = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY ||
  import.meta.env.VITE_SUPABASE_ANON_KEY;

const supabase = url && key ? createClient(url, key) : null;
const ADMIN_EMAIL = 'rdansena25@gmail.com';

const DEPTS = [
  'Structure',
  'Pressure Part',
  'Non Pressure Part',
  'Rotary Part',
  'Bunker',
  'LP Piping',
  'Trim Piping'
];

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

const blank = {
  work_date: new Date().toISOString().slice(0, 10),
  department: 'Structure',
  pg: '',
  ma: '',
  du: '',
  description: '',
  quantity: 0,
  unit_weight: 0,
  pre_assembly: 0,
  placement: 0,
  alignment: 0,
  bolting: 0,
  remark: ''
};

function num(v) {
  const n = Number(
    String(v ?? 0)
      .replace(/,/g, '')
      .replace(/[₹\s]/g, '')
      .replace(/[^\d.-]/g, '')
  );
  return Number.isFinite(n) ? n : 0;
}

function weight(q, u) {
  return num(q) * num(u);
}

function norm(v) {
  return String(v ?? '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function deptName(v) {
  const n = norm(v);
  return DEPTS.find(d => norm(d) === n) || '';
}

function excelDate(v) {
  if (v instanceof Date && !isNaN(v)) {
    return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
  }

  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) {
      return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
    }
  }

  const s = String(v ?? '').trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  // DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY
  let m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);

  if (m) {
    let y = Number(m[3]);
    if (y < 100) y += 2000;

    return `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  }

  // 01 NOV 2026
  m = s.match(/^(\d{1,2})[\s-]+([A-Za-z]{3,})[\s-]+(\d{2,4})$/);

  if (m) {
    const months = [
      'jan', 'feb', 'mar', 'apr', 'may', 'jun',
      'jul', 'aug', 'sep', 'oct', 'nov', 'dec'
    ];

    const mo = months.findIndex(x =>
      m[2].toLowerCase().startsWith(x)
    ) + 1;

    let y = Number(m[3]);
    if (y < 100) y += 2000;

    if (mo) {
      return `${y}-${String(mo).padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    }
  }

  return '';
}

function parseBillText(text) {
  return String(text || '')
    .split(/\r?\n/)
    .filter(line => line.trim())
    .map(line =>
      line.split('\t').length > 1
        ? line.split('\t')
        : line.split('|').map(x => x.trim())
    );
}

function billRows(bill) {
  return bill?.abstract_data?.rows || [];
}

function money(v) {
  return '₹' + num(v).toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

/*
 * Excel Abstract से Bill की जानकारी पहचानने का प्रयास।
 * पूरी Original Abstract Table भी अलग से सुरक्षित रहती है।
 */
function normalizeRaBillNo(value) {
  const text = String(value ?? '').replace(/\s+/g, ' ').trim();
  const match = text.match(/(?:RA\s*(?:BILL)?\s*[-/#:]?\s*)(\d{1,4})/i) || (/^\d{1,4}$/.test(text) ? [text, text] : null);
  if (!match) return text;
  return `RA BILL -${String(match[1]).padStart(2, '0')}`;
}

function extractBillDetails(rows, type) {
  const matrix = (rows || []).map(row => Array.isArray(row) ? row : Object.values(row || {}));
  const flat = matrix.flat().map(v => String(v ?? '').trim());
  const findInline = (patterns) => {
    for (const row of matrix) for (const cell of row) {
      const value = String(cell ?? '').replace(/\s+/g, ' ').trim();
      for (const pattern of patterns) {
        const re = new RegExp(`(?:^|\\b)${pattern}\\s*[:：-]\\s*(.+)$`, 'i');
        const m = value.match(re);
        if (m && m[1].trim()) return m[1].trim().replace(/^[-–—]\s*/, '');
      }
    }
    return '';
  };
  const findRow = (pattern) => matrix.find(row => row.some(c => pattern.test(String(c ?? '')))) || [];
  const cell = (row, index) => row?.[index] ?? '';
  const findLabelValue = (labels) => {
    const wanted = labels.map(norm);
    for (let r = 0; r < matrix.length; r++) {
      const row = matrix[r] || [];
      for (let c = 0; c < row.length; c++) {
        const v = norm(row[c]);
        if (wanted.some(w => v === w || v.startsWith(w))) {
          for (let k = c + 1; k < row.length; k++) if (String(row[k] ?? '').trim()) return row[k];
          for (let rr = r + 1; rr < Math.min(matrix.length, r + 3); rr++) {
            const next = (matrix[rr] || []).find(x => String(x ?? '').trim());
            if (next !== undefined) return next;
          }
        }
      }
    }
    return '';
  };

  const result = {
    agency_name: '', project_name: '', bill_no: '', bill_date: '', billing_period: '',
    previous_amount: 0, current_amount: 0, cumulative_amount: 0,
    gst_amount: 0, gross_bill: 0, net_payable: 0, total_amount: 0
  };

  if (type === 'vendor') {
    result.agency_name = findInline(['Name of Agency']) || flat.find(v => /P\s*E\s*ERECTORS|PRIVATE LIMITED/i.test(v)) || '';
    result.bill_no = normalizeRaBillNo(findInline(['Bill No', 'RA Bill No']) || findLabelValue(['bill no', 'bill number']));
    result.bill_date = excelDate(findInline(['Date']) || findLabelValue(['bill date', 'date of bill']));
    result.billing_period = findInline(['Billing Period']) || findLabelValue(['billing period']);

    const grossRow = findRow(/gross\s+bill\s+value/i);
    const netRow = findRow(/net\s+payable\s+amount/i);
    const totalRow = findRow(/grand\s+total\s+amount/i);
    result.gross_bill = num(cell(grossRow, 10)) || num(cell(totalRow, 10));
    result.net_payable = num(cell(netRow, 10));
    result.previous_amount = num(cell(totalRow, 9));
    result.current_amount = num(cell(totalRow, 10));
    result.cumulative_amount = num(cell(totalRow, 11));
    result.total_amount = result.gross_bill || result.current_amount || result.cumulative_amount;
  } else {
    result.project_name = String(matrix?.[0]?.[0] || 'BHEL PSER LARA SITE').trim();
    result.bill_no = normalizeRaBillNo(findInline(['BILL NO', 'Bill No', 'RA Bill No']) || findLabelValue(['bill no', 'bill number']));
    result.billing_period = findInline(['Billing Period']) || findLabelValue(['billing period']);
    const customerDateRow = matrix.find(row => row.some(value => norm(value) === 'date')) || [];
    const customerDateCol = customerDateRow.findIndex(value => norm(value) === 'date');
    result.bill_date = excelDate((customerDateCol >= 0 ? customerDateRow[customerDateCol + 1] : '') || findInline(['Date']) || findLabelValue(['bill date', 'date of bill']));
    const totalRow = findRow(/total\s+amount/i);
    // Customer template: amount claimed columns J/K/L = previous/current/cumulative.
    result.previous_amount = num(cell(totalRow, 9));
    result.current_amount = num(cell(totalRow, 10));
    result.cumulative_amount = num(cell(totalRow, 11));
    result.gross_bill = result.current_amount;
    result.total_amount = result.current_amount || result.cumulative_amount;
    result.net_payable = result.current_amount;
    result.project_name = result.project_name || 'BHEL PSER LARA SITE';
  }
  return result;
}

function App() {
  const [session, setSession] = useState(null);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [signup, setSignup] = useState(false);
  const [tab, setTab] = useState('dashboard');

  const [rows, setRows] = useState([]);
  const [bills, setBills] = useState([]);
  const [customerBills, setCustomerBills] = useState([]);
  const [vendorPayments, setVendorPayments] = useState([]);
  const [customerReceipts, setCustomerReceipts] = useState([]);
  const [docs, setDocs] = useState([]);

  const [form, setForm] = useState(blank);
  const [filterMonth, setFilterMonth] = useState(new Date().getMonth());
  const [filterYear, setFilterYear] = useState(new Date().getFullYear());
  const [filterDept, setFilterDept] = useState('All');
  const [notice, setNotice] = useState('');
  const [busy, setBusy] = useState(false);
  const [mobile, setMobile] = useState(false);
  const [stats, setStats] = useState(null);

  const [vendorForm, setVendorForm] = useState({
    agency_name: '',
    bill_no: '',
    bill_date: new Date().toISOString().slice(0, 10),
    billing_period: '',
    total_amount: 0,
    net_payable: 0
  });

  const [customerForm, setCustomerForm] = useState({
    project_name: 'BHEL PSER LARA SITE',
    contract_ref: 'CS-9587-001R-2-SC-NOA-7333-AMDT-01',
    bill_no: '',
    bill_date: new Date().toISOString().slice(0, 10),
    billing_period: '',
    total_amount: 0
  });

  const [vendorAbstract, setVendorAbstract] = useState({
    rows: [],
    sheetName: '',
    extracted: {}
  });

  const [customerAbstract, setCustomerAbstract] = useState({
    rows: [],
    sheetName: '',
    extracted: {}
  });

  const [vendorPaste, setVendorPaste] = useState('');
  const [customerPaste, setCustomerPaste] = useState('');
  const [vendorOriginalFile, setVendorOriginalFile] = useState(null);
  const [customerOriginalFile, setCustomerOriginalFile] = useState(null);
  const [selectedBill, setSelectedBill] = useState(null);
  const [customerUnlocked, setCustomerUnlocked] = useState(false);
  const [customerLockInput, setCustomerLockInput] = useState('');

  const [paymentForm, setPaymentForm] = useState({
    bill_id: '',
    payment_date: new Date().toISOString().slice(0, 10),
    amount: '',
    reference_no: '',
    remarks: ''
  });

  const [receiptForm, setReceiptForm] = useState({
    bill_id: '',
    receipt_date: new Date().toISOString().slice(0, 10),
    amount: '',
    reference_no: '',
    remarks: ''
  });

  useEffect(() => {
    if (!supabase) return;

    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session);
    });

    const {
      data: { subscription }
    } = supabase.auth.onAuthStateChange((_event, currentSession) => {
      setSession(currentSession);
    });

    return () => subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) return;

    loadAll();
    supabase.rpc('record_user_activity');

    const timer = setInterval(() => {
      supabase.rpc('record_user_activity');
    }, 60000);

    return () => clearInterval(timer);
  }, [session]);

  async function loadAll() {
    if (!session || !supabase) return;

    setBusy(true);

    try {
      const results = await Promise.all([
        supabase.from('dpr_entries').select('*').order('work_date', { ascending: false }),
        supabase.from('vendor_bills').select('*').order('created_at', { ascending: false }),
        supabase.from('customer_bill_abstracts').select('*').order('created_at', { ascending: false }),
        supabase.from('drawings').select('*').order('created_at', { ascending: false }),
        supabase.from('vendor_bill_payments').select('*').order('payment_date', { ascending: false }),
        supabase.from('customer_bill_receipts').select('*').order('receipt_date', { ascending: false })
      ]);

      const [d, v, c, f, p, r] = results;

      if (d.error) setNotice('DPR Load Error: ' + d.error.message);
      else setRows(d.data || []);

      if (v.error) setNotice('Vendor Bill Load Error: ' + v.error.message);
      else setBills(v.data || []);

      if (c.error) setNotice('Customer Bill Load Error: ' + c.error.message);
      else setCustomerBills(c.data || []);

      if (f.error) setNotice('Drawing Load Error: ' + f.error.message);
      else setDocs(f.data || []);

      if (p.error) setNotice('Vendor Payment Load Error: ' + p.error.message);
      else setVendorPayments(p.data || []);

      if (r.error) setNotice('Customer Receipt Load Error: ' + r.error.message);
      else setCustomerReceipts(r.data || []);

      if (session.user.email?.toLowerCase() === ADMIN_EMAIL) {
        const s = await supabase.rpc('get_admin_user_stats');
        if (!s.error) setStats(s.data);
      }
    } catch (err) {
      setNotice('Load Error: ' + err.message);
    } finally {
      setBusy(false);
    }
  }

  async function login(e) {
    e.preventDefault();

    if (!supabase) return;

    setBusy(true);

    try {
      const action = signup
        ? supabase.auth.signUp({ email, password })
        : supabase.auth.signInWithPassword({ email, password });

      const { error } = await action;

      setNotice(error
        ? error.message
        : signup
          ? 'Account created. Email confirmation may be required before login.'
          : 'Login successful'
      );
    } catch (err) {
      setNotice(err.message);
    } finally {
      setBusy(false);
    }
  }

  async function adminToken() {
    const p = window.prompt('Admin Password डालें:');
    if (!p) return null;

    const { data, error } = await supabase.rpc(
      'verify_admin_password',
      { p_password: p }
    );

    if (error) {
      setNotice('Admin Password Error: ' + error.message);
      return null;
    }

    if (!data) {
      setNotice('Admin Password गलत है।');
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

    const { error } = await supabase.rpc('set_admin_password', {
      p_password: p
    });

    setNotice(
      error
        ? error.message
        : 'Admin Password सफलतापूर्वक सेट हो गया।'
    );
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
      work_date: form.work_date,
      department: form.department,
      pg: String(form.pg).trim(),
      ma: String(form.ma).trim(),
      du: String(form.du).trim(),
      description: String(form.description || ''),
      quantity: num(form.quantity),
      unit_weight: num(form.unit_weight),
      weight: weight(form.quantity, form.unit_weight),
      pre_assembly: num(form.pre_assembly),
      placement: num(form.placement),
      alignment: num(form.alignment),
      bolting: num(form.bolting),
      remark: String(form.remark || '')
    };

    const { error } = await supabase.rpc('save_dpr_entry', {
      p_token: token,
      p_payload: payload
    });

    setNotice(error
      ? error.message
      : 'DPR saved / updated successfully'
    );

    if (!error) {
      setForm({
        ...blank,
        work_date: form.work_date,
        department: form.department
      });
      await loadAll();
    }
  }

  async function delDpr(id) {
    if (!window.confirm('इस DPR को Delete करना है?')) return;

    const token = await adminToken();
    if (!token) return;

    const { error } = await supabase.rpc('delete_dpr_entry', {
      p_token: token,
      p_id: id
    });

    setNotice(error ? error.message : 'DPR deleted');

    if (!error) await loadAll();
  }

  const monthRows = useMemo(() => rows.filter(r => {
    const d = new Date(r.work_date + 'T00:00:00');
    return d.getMonth() === filterMonth &&
      d.getFullYear() === filterYear;
  }), [rows, filterMonth, filterYear]);

  const visibleRows = useMemo(() =>
    filterDept === 'All'
      ? monthRows
      : monthRows.filter(r => r.department === filterDept),
    [monthRows, filterDept]
  );

  const totals = useMemo(() => rows.reduce((a, r) => ({
    qty: a.qty + num(r.quantity),
    weight: a.weight + num(r.weight)
  }), { qty: 0, weight: 0 }), [rows]);

  const vendorMonthBills = bills.filter(b => {
    const d = b.bill_date ? new Date(b.bill_date + 'T00:00:00') : null;
    return d &&
      d.getMonth() === filterMonth &&
      d.getFullYear() === filterYear;
  });

  const customerMonthBills = customerBills.filter(b => {
    const d = b.bill_date ? new Date(b.bill_date + 'T00:00:00') : null;
    return d &&
      d.getMonth() === filterMonth &&
      d.getFullYear() === filterYear;
  });

  function paidFor(id) {
    return vendorPayments
      .filter(p => p.bill_id === id)
      .reduce((s, p) => s + num(p.amount), 0);
  }

  function receivedFor(id) {
    return customerReceipts
      .filter(r => r.bill_id === id)
      .reduce((s, r) => s + num(r.amount), 0);
  }

  function vendorOutstanding(b) {
    return num(b.net_payable || b.total_amount) - paidFor(b.id);
  }

  function customerOutstanding(b) {
    return num(b.total_amount) - receivedFor(b.id);
  }

  const vendorTotals = vendorMonthBills.reduce((a, b) => ({
    billed: a.billed + num(b.net_payable || b.total_amount),
    paid: a.paid + paidFor(b.id)
  }), { billed: 0, paid: 0 });

  const customerTotals = customerMonthBills.reduce((a, b) => ({
    billed: a.billed + num(b.total_amount),
    received: a.received + receivedFor(b.id)
  }), { billed: 0, received: 0 });

  function exportRows(data, name) {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, 'Records');
    XLSX.writeFile(wb, name + '.xlsx');
  }

  function exportDprWorkbook() {
    const wb = XLSX.utils.book_new();

    const headers = [
      'Date', 'PG', 'MA', 'DU', 'Description',
      'Quantity', 'Unit Weight', 'Total Weight',
      'Pre-Assembly Qty', 'Placement Qty',
      'Alignment Qty', 'Bolting Qty', 'Remarks'
    ];

    for (const dept of DEPTS) {
      const data = rows.filter(r => r.department === dept);

      const aoa = [
        ['BOILER CONSTRUCTION MANAGEMENT'],
        [`DAILY PROGRESS REPORT - ${dept}`],
        ['Generated: ' + new Date().toLocaleDateString('en-IN')],
        headers
      ];

      data.forEach(r => aoa.push([
        r.work_date,
        r.pg,
        r.ma,
        r.du,
        r.description,
        num(r.quantity),
        num(r.unit_weight),
        num(r.weight),
        num(r.pre_assembly),
        num(r.placement),
        num(r.alignment),
        num(r.bolting),
        r.remark || ''
      ]));

      const ws = XLSX.utils.aoa_to_sheet(aoa);

      ws['!cols'] = [
        15, 12, 12, 12, 30, 14, 14,
        16, 17, 15, 15, 14, 28
      ].map(w => ({ wch: w }));

      ws['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 12 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 12 } },
        { s: { r: 2, c: 0 }, e: { r: 2, c: 12 } }
      ];

      XLSX.utils.book_append_sheet(wb, ws, dept);
    }

    XLSX.writeFile(wb, 'Boiler_Construction_DPR.xlsx');
  }

  async function importDpr(e) {
    const file = e.target.files?.[0];
    if (!file) return;

    const token = await adminToken();

    if (!token) {
      e.target.value = '';
      return;
    }

    setBusy(true);

    try {
      const wb = XLSX.read(await file.arrayBuffer(), {
        cellDates: true
      });

      const mapped = [];
      let skipped = 0;

      for (const sheetName of wb.SheetNames) {
        const matrix = XLSX.utils.sheet_to_json(
          wb.Sheets[sheetName],
          { header: 1, defval: '', raw: true }
        );

        let hi = -1;
        let headers = [];

        for (let i = 0; i < Math.min(matrix.length, 25); i++) {
          const h = matrix[i].map(norm);

          if (
            h.includes('pg') &&
            h.includes('ma') &&
            h.includes('du') &&
            h.some(x => ['date', 'workdate'].includes(x))
          ) {
            hi = i;
            headers = h;
            break;
          }
        }

        if (hi < 0) continue;

        const sheetDept = deptName(sheetName);

        const get = (r, aliases) => {
          const i = headers.findIndex(h => aliases.includes(h));
          return i < 0 ? '' : r[i];
        };

        for (const r of matrix.slice(hi + 1)) {
          if (!r.some(v => String(v ?? '').trim())) continue;

          const date = excelDate(get(r, ['date', 'workdate']));
          const department =
            sheetDept || deptName(get(r, ['department', 'dept']));

          const pg = get(r, ['pg']);
          const ma = get(r, ['ma']);
          const du = get(r, ['du']);

          if (!date || !DEPTS.includes(department) || !pg || !ma || !du) {
            skipped++;
            continue;
          }

          const quantity = num(get(r, ['quantity', 'qty']));
          const unit_weight = num(get(r, ['unitweight']));

          mapped.push({
            work_date: date,
            department,
            pg: String(pg).trim(),
            ma: String(ma).trim(),
            du: String(du).trim(),
            description: String(
              get(r, ['description', 'itemdescription']) || ''
            ),
            quantity,
            unit_weight,
            weight: quantity * unit_weight,
            pre_assembly: num(get(r, ['preassemblyqty', 'preassembly'])),
            placement: num(get(r, ['placementqty', 'placement'])),
            alignment: num(get(r, ['alignmentqty', 'alignment'])),
            bolting: num(get(r, ['boltingqty', 'bolting'])),
            remark: String(
              get(r, ['remarks', 'remark', 'status']) || ''
            )
          });
        }
      }

      if (!mapped.length) {
        throw new Error(
          'Excel में सही Date, PG, MA, DU और Department headers नहीं मिले।'
        );
      }

      const { data, error } = await supabase.rpc('import_dpr_entries', {
        p_token: token,
        p_rows: mapped
      });

      if (error) throw error;

      setNotice(`${data} DPR rows imported/updated. Skipped: ${skipped}`);
      await loadAll();
    } catch (err) {
      setNotice('Excel Import Error: ' + err.message);
    } finally {
      setBusy(false);
      e.target.value = '';
    }
  }

  function applyExtractedBill(type, extracted, data) {
    if (type === 'vendor') {
      setVendorAbstract({
        ...data,
        extracted
      });

      setVendorForm(old => ({
        ...old,
        agency_name: extracted.agency_name || old.agency_name,
        bill_no: extracted.bill_no || old.bill_no,
        bill_date: extracted.bill_date || old.bill_date,
        billing_period: extracted.billing_period || old.billing_period,
        total_amount: extracted.gross_bill ||
          extracted.current_amount ||
          extracted.total_amount ||
          old.total_amount,
        net_payable: extracted.net_payable || old.net_payable
      }));
    } else {
      setCustomerAbstract({
        ...data,
        extracted
      });

      setCustomerForm(old => ({
        ...old,
        project_name: extracted.project_name || old.project_name,
        bill_no: extracted.bill_no || old.bill_no,
        bill_date: extracted.bill_date || old.bill_date,
        billing_period: extracted.billing_period || old.billing_period,
        total_amount: extracted.current_amount ||
          extracted.gross_bill ||
          extracted.total_amount ||
          old.total_amount
      }));
    }
  }

  function readBillFile(e, type) {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = event => {
      try {
        const wb = XLSX.read(event.target.result, {
          type: 'array', cellDates: true, cellStyles: true, cellFormula: true
        });

        const sheetName = wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];

        const matrix = XLSX.utils.sheet_to_json(ws, {
          header: 1,
          defval: '',
          raw: false
        });

        const data = {
          rows: matrix,
          sheetName,
          merges: ws['!merges'] || [],
          cols: (ws['!cols'] || []).map(c => ({ wch: c.wch || 12 })),
          ref: ws['!ref'] || '',
          originalFileName: file.name,
          originalFileSize: file.size,
          originalFileType: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        };

        if (type === 'vendor') setVendorOriginalFile(file);
        else setCustomerOriginalFile(file);

        const extracted = extractBillDetails(matrix, type);

        applyExtractedBill(type, extracted, data);

        setNotice(
          `${file.name} पढ़ ली गई। उपलब्ध Bill Details अपने-आप भरी गई हैं। Save करने से पहले जाँच लें।`
        );
      } catch (err) {
        setNotice('Excel पढ़ने में Error: ' + err.message);
      }

      e.target.value = '';
    };

    reader.readAsArrayBuffer(file);
  }

  function handleBillPaste(value, type) {
    if (type === 'vendor') setVendorPaste(value);
    else setCustomerPaste(value);

    if (!value.trim()) return;

    const matrix = parseBillText(value);
    const extracted = extractBillDetails(matrix, type);

    const data = {
      rows: matrix,
      sheetName: 'Pasted Abstract'
    };

    applyExtractedBill(type, extracted, data);
  }

  async function uploadBillOriginal(file, type, billNo) {
    if (!file) return null;
    const safeName = String(file.name || 'bill.xlsx').replace(/[^a-zA-Z0-9._-]/g, '_');
    const folder = type === 'vendor' ? '01_Vendor_Bills' : '02_Customer_Bills';
    const path = `${session.user.id}/Master_Records/${folder}/${Date.now()}_${safeName}`;
    const { data, error } = await supabase.storage.from('drawings').upload(path, file, {
      contentType: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      upsert: false
    });
    if (error) throw new Error('Original Excel Archive Upload Error: ' + error.message);
    return { storage_path: data.path || path, file_name: file.name, file_size: file.size, bill_no: billNo || '' };
  }

  async function saveVendorBill(e) {
    e.preventDefault();

    let data = vendorAbstract;

    if (vendorPaste.trim()) {
      const matrix = parseBillText(vendorPaste);
      data = {
        rows: matrix,
        sheetName: 'Pasted Abstract',
        extracted: extractBillDetails(matrix, 'vendor')
      };
    }

    if (!data.rows?.length) {
      setNotice('पहले Vendor Abstract की Excel Import या Copy-Paste करें।');
      return;
    }

    const token = await adminToken();
    if (!token) return;

    setBusy(true);

    let originalFile = null;
    try { originalFile = await uploadBillOriginal(vendorOriginalFile, 'vendor', normalizeRaBillNo(vendorForm.bill_no)); }
    catch (err) { setBusy(false); setNotice(err.message); return; }

    const payload = {
      ...vendorForm, bill_no: normalizeRaBillNo(vendorForm.bill_no), total_amount: num(vendorForm.total_amount),
      net_payable: num(vendorForm.net_payable),
      abstract_data: { ...data, archive: originalFile, archive_folder: '01_Vendor_Bills' }
    };

    const { error } = await supabase.rpc('save_vendor_bill', {
      p_token: token,
      p_payload: payload
    });

    setBusy(false);

    if (error) {
      setNotice('Vendor Bill Error: ' + error.message);
    } else {
      setNotice('Vendor Bill सफलतापूर्वक सेव / अपडेट हो गया।');

      setVendorForm({
        agency_name: '',
        bill_no: '',
        bill_date: new Date().toISOString().slice(0, 10),
        billing_period: '',
        total_amount: 0,
        net_payable: 0
      });

      setVendorAbstract({
        rows: [],
        sheetName: '',
        extracted: {}
      });

      setVendorPaste('');
      setVendorOriginalFile(null);
      await loadAll();
    }
  }

  async function saveCustomerBill(e) {
    e.preventDefault();

    let data = customerAbstract;

    if (customerPaste.trim()) {
      const matrix = parseBillText(customerPaste);
      data = {
        rows: matrix,
        sheetName: 'Pasted Abstract',
        extracted: extractBillDetails(matrix, 'customer')
      };
    }

    if (!data.rows?.length) {
      setNotice('पहले Customer Abstract की Excel Import या Copy-Paste करें।');
      return;
    }

    const token = await adminToken();
    if (!token) return;

    setBusy(true);

    let originalFile = null;
    try { originalFile = await uploadBillOriginal(customerOriginalFile, 'customer', normalizeRaBillNo(customerForm.bill_no)); }
    catch (err) { setBusy(false); setNotice(err.message); return; }

    const payload = {
      ...customerForm, bill_no: normalizeRaBillNo(customerForm.bill_no), total_amount: num(customerForm.total_amount),
      abstract_data: { ...data, archive: originalFile, archive_folder: '02_Customer_Bills' }
    };

    const { error } = await supabase.rpc('save_customer_bill', {
      p_token: token,
      p_payload: payload
    });

    setBusy(false);

    if (error) {
      setNotice('Customer Bill Error: ' + error.message);
    } else {
      setNotice('Customer Bill सफलतापूर्वक सेव / अपडेट हो गया।');

      setCustomerForm({
        project_name: 'BHEL PSER LARA SITE',
        contract_ref: 'CS-9587-001R-2-SC-NOA-7333-AMDT-01',
        bill_no: '',
        bill_date: new Date().toISOString().slice(0, 10),
        billing_period: '',
        total_amount: 0
      });

      setCustomerAbstract({
        rows: [],
        sheetName: '',
        extracted: {}
      });

      setCustomerPaste('');
      setCustomerOriginalFile(null);
      await loadAll();
    }
  }

  async function addVendorPayment(e) {
    e.preventDefault();

    if (!paymentForm.bill_id || num(paymentForm.amount) <= 0) {
      setNotice('Vendor Bill और Payment Amount चुनना जरूरी है।');
      return;
    }

    const token = await adminToken();
    if (!token) return;

    const { error } = await supabase.rpc('add_vendor_payment', {
      p_token: token,
      p_payload: {
        ...paymentForm,
        amount: num(paymentForm.amount)
      }
    });

    if (error) {
      setNotice('Payment Error: ' + error.message);
    } else {
      setNotice('Vendor Payment दर्ज हो गया।');

      setPaymentForm({
        ...paymentForm,
        amount: '',
        reference_no: '',
        remarks: ''
      });

      await loadAll();
    }
  }

  async function addCustomerReceipt(e) {
    e.preventDefault();

    if (!receiptForm.bill_id || num(receiptForm.amount) <= 0) {
      setNotice('Customer Bill और Receipt Amount चुनना जरूरी है।');
      return;
    }

    const token = await adminToken();
    if (!token) return;

    const { error } = await supabase.rpc('add_customer_receipt', {
      p_token: token,
      p_payload: {
        ...receiptForm,
        amount: num(receiptForm.amount)
      }
    });

    if (error) {
      setNotice('Receipt Error: ' + error.message);
    } else {
      setNotice('Customer Receipt दर्ज हो गया।');

      setReceiptForm({
        ...receiptForm,
        amount: '',
        reference_no: '',
        remarks: ''
      });

      await loadAll();
    }
  }

  async function uploadDrawing(e) {
    const file = e.target.files?.[0];
    if (!file) return;

    setBusy(true);

    const path = `${session.user.id}/${Date.now()}-${file.name}`;

    const up = await supabase.storage
      .from('drawings')
      .upload(path, file);

    if (up.error) {
      setNotice(up.error.message);
      setBusy(false);
      e.target.value = '';
      return;
    }

    const { error } = await supabase.rpc('create_drawing_record', {
      p_payload: {
        file_name: file.name,
        storage_path: path,
        department: form.department || 'General'
      }
    });

    setNotice(
      error
        ? 'Drawing record error: ' + error.message
        : 'Drawing uploaded successfully'
    );

    if (error) {
      await supabase.storage.from('drawings').remove([path]);
    } else {
      await loadAll();
    }

    setBusy(false);
    e.target.value = '';
  }

  function printAbstract() {
    window.print();
  }

  if (!supabase) return <Setup />;

  if (!session) {
    return (
      <div className="login-wrap">
        <form className="login-card" onSubmit={login}>
          <div className="logo">
            <Factory size={34} />
          </div>

          <h1>Boiler Management</h1>
          <p>Construction Management System</p>

          <label>Email</label>
          <input
            type="email"
            value={email}
            onChange={e => setEmail(e.target.value)}
            required
            placeholder="Email"
          />

          <label>Password</label>
          <input
            type="password"
            value={password}
            onChange={e => setPassword(e.target.value)}
            required
            placeholder="Password"
          />

          <button className="primary full" disabled={busy}>
            <LockKeyhole size={17} />
            {signup ? 'Create Account' : 'Secure Login'}
          </button>

          <button
            type="button"
            className="link"
            onClick={() => setSignup(!signup)}
          >
            {signup
              ? 'पहले से Account है? Login करें'
              : 'नया User ID बनाएँ / Sign Up'}
          </button>

          {notice && <small>{notice}</small>}
        </form>
      </div>
    );
  }

  const nav = [
    ['dashboard', 'Dashboard', LayoutDashboard],
    ['master', 'Master Dashboard', LayoutDashboard],
    ['records', 'Master Records', FolderOpen],
    ['dpr', 'DPR Register', ClipboardList],
    ['ra', 'Vendor Bills', Receipt],
    ['customer', 'Customer Bills', FileText],
    ['drawings', 'Drawing Folder', FolderOpen]
  ];

  return (
    <div className="app">
      <aside className={mobile ? 'open' : ''}>
        <div className="brand">
          <Factory />
          <span>
            BOILER
            <br />
            <b>MANAGEMENT</b>
          </span>

          <button
            className="mobile-close"
            onClick={() => setMobile(false)}
          >
            <X />
          </button>
        </div>

        <div className="nav-label">WORKSPACE</div>

        {nav.map(([id, label, Icon]) => (
          <button
            key={id}
            className={'nav ' + (tab === id ? 'active' : '')}
            onClick={() => {
              setTab(id);
              setMobile(false);
            }}
          >
            <Icon size={18} />
            {label}
          </button>
        ))}

        <div className="side-bottom">
          <span>{session.user.email}</span>

          <button
            className="nav"
            onClick={() => supabase.auth.signOut()}
          >
            <LogOut size={18} />
            Logout
          </button>
        </div>
      </aside>

      <main>
        <header>
          <button
            className="menu"
            onClick={() => setMobile(!mobile)}
          >
            <Menu />
          </button>

          <div>
            <b>{nav.find(n => n[0] === tab)?.[1] || 'Dashboard'}</b>
            <small>Boiler Construction Project Workspace</small>
          </div>

          <button
            className="icon-btn"
            onClick={loadAll}
            disabled={busy}
          >
            <RefreshCw size={18} />
          </button>
        </header>

        <div className="content">
          {notice && (
            <div
              className="notice"
              onClick={() => setNotice('')}
            >
              {notice}　×
            </div>
          )}

          {session.user.email?.toLowerCase() === ADMIN_EMAIL && (
            <div className="panel">
              <button
                className="secondary"
                onClick={setAdminPassword}
              >
                <LockKeyhole size={16} />
                Set / Change Admin Password
              </button>

              {stats && (
                <p>
                  <Users size={16} />
                  Registered Users: <b>{stats.total_users}</b>
                  {' '} | Active Users (last 5 min):{' '}
                  <b>{stats.active_users}</b>
                </p>
              )}
            </div>
          )}

          {tab === 'dashboard' && (
            <>
              <div className="welcome">
                <div>
                  <span className="eyebrow">PROJECT OVERVIEW</span>
                  <h1>Master Dashboard</h1>
                  <p>All construction records in one place.</p>
                </div>

                <button
                  onClick={() => setTab('dpr')}
                  className="primary"
                >
                  <Plus size={17} />
                  Add DPR
                </button>
              </div>

              <div className="panel">
                <div className="filters">
                  <label>
                    Month
                    <select
                      value={filterMonth}
                      onChange={e => setFilterMonth(+e.target.value)}
                    >
                      {MONTHS.map((m, i) => (
                        <option key={m} value={i}>{m}</option>
                      ))}
                    </select>
                  </label>

                  <label>
                    Year
                    <input
                      type="number"
                      value={filterYear}
                      onChange={e => setFilterYear(+e.target.value)}
                    />
                  </label>
                </div>
              </div>

              <div className="stats">
                <Stat title="DPR Entries" value={rows.length} />
                <Stat title="Total Quantity" value={totals.qty.toLocaleString()} />
                <Stat title="Total Weight" value={totals.weight.toLocaleString()} />
                <Stat title="Vendor Bills" value={bills.length} />
                <Stat title="Customer Bills" value={customerBills.length} />
                <Stat title="Drawings" value={docs.length} />
              </div>

              <div className="stats">
                <Stat title="Vendor Billed" value={money(vendorTotals.billed)} />
                <Stat title="Vendor Paid" value={money(vendorTotals.paid)} />
                <Stat title="Vendor Outstanding" value={money(vendorTotals.billed - vendorTotals.paid)} />
                <Stat title="Customer Billed" value={money(customerTotals.billed)} />
                <Stat title="Customer Received" value={money(customerTotals.received)} />
                <Stat title="Customer Outstanding" value={money(customerTotals.billed - customerTotals.received)} />
              </div>

              <div className="panel">
                <h2>{MONTHS[filterMonth]} {filterYear} - Monthly Abstract Summary</h2>
                <p>Vendor Bills: {vendorMonthBills.length} | Customer Bills: {customerMonthBills.length}</p>
                <p>Vendor Outstanding: <b>{money(vendorTotals.billed - vendorTotals.paid)}</b></p>
                <p>Customer Outstanding: <b>{money(customerTotals.billed - customerTotals.received)}</b></p>
              </div>

              <div className="panel">
                <div className="panel-head">
                  <h2>Department Progress</h2>
                  <span>All-time quantity & weight</span>
                </div>

                <div className="dept-grid">
                  {DEPTS.map(d => {
                    const a = rows.filter(r => r.department === d);

                    return (
                      <div className="dept" key={d}>
                        <b>{d}</b>
                        <strong>
                          {a.reduce((s, r) => s + num(r.quantity), 0).toLocaleString()}
                        </strong>
                        <small>
                          {a.length} entries ·{' '}
                          {a.reduce((s, r) => s + num(r.weight), 0).toLocaleString()} weight
                        </small>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="panel">
                <h2>Recent DPR Entries</h2>
                <DprTable rows={rows.slice(0, 8)} onDelete={delDpr} />
              </div>
            </>
          )}

          {tab === 'master' && (() => {
            const allVendorBilled = bills.reduce((s, b) => s + num(b.net_payable || b.total_amount), 0);
            const allVendorPaid = vendorPayments.reduce((s, p) => s + num(p.amount), 0);
            const allCustomerBilled = customerBills.reduce((s, b) => s + num(b.total_amount), 0);
            const allCustomerReceived = customerReceipts.reduce((s, r) => s + num(r.amount), 0);
            const difference = allCustomerBilled - allVendorBilled;
            const dprWeight = rows.reduce((s, r) => s + num(r.weight), 0);
            return <>
              <div className="welcome"><div><span className="eyebrow">COMPLETE FINANCIAL & WORK SUMMARY</span><h1>Master Dashboard</h1><p>सभी Vendor, Customer, Payment, Receipt और DPR का कुल हिसाब</p></div><button className="secondary" onClick={() => setTab('records')}><FolderOpen size={16}/> Master Records खोलें</button></div>
              <div className="stats">
                <Stat title="Total Vendor Bills" value={bills.length}/><Stat title="Vendor Bill Amount" value={money(allVendorBilled)}/><Stat title="Vendor Paid" value={money(allVendorPaid)}/><Stat title="Vendor Balance" value={money(allVendorBilled-allVendorPaid)}/>
                <Stat title="Total Customer Bills" value={customerBills.length}/><Stat title="Customer Bill Amount" value={money(allCustomerBilled)}/><Stat title="Customer Received" value={money(allCustomerReceived)}/><Stat title="Customer Balance" value={money(allCustomerBilled-allCustomerReceived)}/>
              </div>
              <div className="panel"><h2>Customer Bill बनाम Vendor Bill</h2><div className="stats"><Stat title="Customer - Vendor Difference" value={money(difference)}/><Stat title="Total DPR Entries" value={rows.length}/><Stat title="Total DPR Weight" value={dprWeight.toLocaleString('en-IN')}/><Stat title="Drawings Saved" value={docs.length}/></div><p>{difference >= 0 ? 'Customer billing, Vendor billing से अधिक है।' : 'Vendor billing, Customer billing से अधिक है।'} यह अंतर बिल किए गए कुल मान पर आधारित है; वास्तविक लाभ/लागत नहीं।</p></div>
              <div className="panel"><h2>Bill-wise Comparison</h2><div className="table-wrap"><table><thead><tr><th>Type</th><th>Bill No.</th><th>Date</th><th>Bill Amount</th><th>Paid / Received</th><th>Balance</th></tr></thead><tbody>
                {bills.map(b => <tr key={'v'+b.id}><td>Vendor</td><td>{b.bill_no}</td><td>{b.bill_date}</td><td>{money(b.net_payable || b.total_amount)}</td><td>{money(paidFor(b.id))}</td><td>{money(vendorOutstanding(b))}</td></tr>)}
                {customerBills.map(b => <tr key={'c'+b.id}><td>Customer</td><td>{b.bill_no}</td><td>{b.bill_date}</td><td>{money(b.total_amount)}</td><td>{money(receivedFor(b.id))}</td><td>{money(customerOutstanding(b))}</td></tr>)}
              </tbody></table></div></div>
            </>;
          })()}

          {tab === 'records' && <>
            <div className="welcome"><div><span className="eyebrow">PERMANENT RECORD ARCHIVE</span><h1>Master Records</h1><p>सभी पुराने और नए रिकॉर्ड अलग-अलग सेक्शन में सुरक्षित</p></div><button className="secondary" onClick={loadAll}><RefreshCw size={16}/> Refresh Records</button></div>
            <div className="stats"><Stat title="01 Vendor Bills" value={bills.length}/><Stat title="02 Customer Bills" value={customerBills.length}/><Stat title="03 DPR Records" value={rows.length}/><Stat title="04 Drawings" value={docs.length}/><Stat title="05 Vendor Payments" value={vendorPayments.length}/><Stat title="06 Customer Receipts" value={customerReceipts.length}/></div>
            <div className="panel"><h2>01 · Vendor Bills Folder</h2><BillTable bills={bills} type="vendor" paidFor={paidFor} outstanding={vendorOutstanding} onView={setSelectedBill}/></div>
            <div className="panel"><h2>02 · Customer Bills Folder</h2>{customerUnlocked ? <BillTable bills={customerBills} type="customer" paidFor={receivedFor} outstanding={customerOutstanding} onView={setSelectedBill}/> : <p>Customer Bill सुरक्षित है। इसे देखने के लिए Customer Bills खोलकर पासवर्ड डालें।</p>}</div>
            <div className="panel"><h2>03 · DPR Records Folder</h2><DprTable rows={rows.slice(0,100)} onDelete={delDpr}/><p>दिखाए गए नवीनतम 100 रिकॉर्ड। कुल रिकॉर्ड: {rows.length}</p></div>
            <div className="panel"><h2>04 · Drawings Folder</h2><DataTable rows={docs} columns={['file_name','department','storage_path','created_at']}/></div>
            <div className="panel"><h2>05 · Vendor Payment Records</h2><DataTable rows={vendorPayments} columns={['payment_date','bill_id','amount','reference_no','remarks']}/></div>
            <div className="panel"><h2>06 · Customer Receipt Records</h2><DataTable rows={customerReceipts} columns={['receipt_date','bill_id','amount','reference_no','remarks']}/></div>
            <div className="panel"><b>Archive सुरक्षा:</b> नया Excel अलग unique filename से Storage के drawings bucket में Master_Records folder के अंतर्गत रखा जाता है। Existing bill/DPR/drawing records को इस स्क्रीन से delete नहीं किया जाता।</div>
          </>}

          {tab === 'dpr' && (
            <>
              <div className="welcome">
                <div>
                  <span className="eyebrow">DAILY PROGRESS REPORT</span>
                  <h1>DPR Register</h1>
                </div>

                <div className="actions">
                  <label className="secondary upload">
                    <Upload size={16} />
                    Import Excel
                    <input
                      type="file"
                      accept=".xlsx,.xls,.csv"
                      onChange={importDpr}
                    />
                  </label>

                  <button
                    className="secondary"
                    onClick={exportDprWorkbook}
                  >
                    <Download size={16} />
                    Export Excel
                  </button>
                </div>
              </div>

              <div className="panel">
                <h2>Add Daily Progress</h2>

                <form className="entry-grid" onSubmit={saveDpr}>
                  <Field label="Date">
                    <input
                      type="date"
                      value={form.work_date}
                      onChange={e => setForm({ ...form, work_date: e.target.value })}
                      required
                    />
                  </Field>

                  <Field label="Department">
                    <select
                      value={form.department}
                      onChange={e => setForm({ ...form, department: e.target.value })}
                    >
                      {DEPTS.map(d => <option key={d}>{d}</option>)}
                    </select>
                  </Field>

                  {[
                    'pg', 'ma', 'du', 'description',
                    'quantity', 'unit_weight', 'pre_assembly',
                    'placement', 'alignment', 'bolting', 'remark'
                  ].map(k => (
                    <Field
                      label={k.replaceAll('_', ' ').toUpperCase()}
                      key={k}
                    >
                      <input
                        type={[
                          'quantity', 'unit_weight', 'pre_assembly',
                          'placement', 'alignment', 'bolting'
                        ].includes(k) ? 'number' : 'text'}
                        step="any"
                        value={form[k]}
                        onChange={e => setForm({
                          ...form,
                          [k]: e.target.value
                        })}
                      />
                    </Field>
                  ))}

                  <Field label="TOTAL WEIGHT (AUTO)">
                    <input
                      value={weight(form.quantity, form.unit_weight)}
                      readOnly
                    />
                  </Field>

                  <div className="form-end">
                    <button className="primary" disabled={busy}>
                      <Plus size={16} />
                      Save / Update DPR
                    </button>
                  </div>
                </form>
              </div>

              <div className="panel">
                <div className="panel-head">
                  <h2>DPR Register</h2>

                  <div className="filters">
                    <select
                      value={filterDept}
                      onChange={e => setFilterDept(e.target.value)}
                    >
                      <option>All</option>
                      {DEPTS.map(d => <option key={d}>{d}</option>)}
                    </select>

                    <select
                      value={filterMonth}
                      onChange={e => setFilterMonth(+e.target.value)}
                    >
                      {MONTHS.map((m, i) => (
                        <option key={m} value={i}>{m}</option>
                      ))}
                    </select>

                    <input
                      type="number"
                      value={filterYear}
                      onChange={e => setFilterYear(+e.target.value)}
                    />
                  </div>
                </div>

                <div className="month-summary">
                  <b>{MONTHS[filterMonth]} {filterYear}</b>
                  <span>{visibleRows.length} entries</span>
                  <span>
                    Qty: {visibleRows.reduce((s, r) => s + num(r.quantity), 0)}
                  </span>
                  <span>
                    Weight: {visibleRows.reduce((s, r) => s + num(r.weight), 0)}
                  </span>
                </div>

                <DprTable rows={visibleRows} onDelete={delDpr} />
              </div>
            </>
          )}

          {tab === 'ra' && (
            <>
              <div className="welcome">
                <div>
                  <span className="eyebrow">VENDOR BILLING</span>
                  <h1>Vendor Bill Management</h1>
                  <p>Import your original Vendor RA Abstract Excel or paste its rows.</p>
                </div>

                <button
                  className="secondary"
                  onClick={() => exportRows(bills.map(b => ({
                    agency_name: b.agency_name,
                    bill_no: b.bill_no,
                    bill_date: b.bill_date,
                    billing_period: b.billing_period,
                    total_amount: b.total_amount,
                    net_payable: b.net_payable
                  })), 'Vendor_Bills')}
                >
                  <Download size={16} />
                  Export Register
                </button>
                <button className="secondary" onClick={() => { setCustomerUnlocked(false); setSelectedBill(null); setCustomerLockInput(''); }}>
                  <LockKeyhole size={16} /> Lock Customer Bills
                </button>
              </div>

              <div className="stats">
                <Stat title="Total Bills" value={vendorMonthBills.length} />
                <Stat title="Total Billed" value={money(vendorTotals.billed)} />
                <Stat title="Paid" value={money(vendorTotals.paid)} />
                <Stat title="Outstanding" value={money(vendorTotals.billed - vendorTotals.paid)} />
              </div>

              <div className="panel">
                <h2>Import Vendor Abstract</h2>
                <p>
                  अपनी मौजूदा Vendor RA Bill Excel फाइल चुनें।
                  Excel से पूरा Abstract Copy करके भी Paste कर सकते हैं।
                </p>

                <label className="secondary upload">
                  <Upload size={16} />
                  Import Vendor Excel
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={e => readBillFile(e, 'vendor')}
                  />
                </label>

                {vendorOriginalFile && <p className="success-text">Original file: {vendorOriginalFile.name} · {Math.ceil(vendorOriginalFile.size / 1024)} KB</p>}

                {vendorAbstract.rows.length > 0 && (
                  <p className="success-text">
                    Excel तैयार: {vendorAbstract.sheetName} · {vendorAbstract.rows.length} rows
                  </p>
                )}

                <Field label="या Excel से पूरा Abstract Copy करके यहाँ Paste करें">
                  <textarea
                    rows="5"
                    value={vendorPaste}
                    onChange={e => handleBillPaste(e.target.value, 'vendor')}
                    placeholder="Excel की पूरी टेबल Copy करके यहाँ Paste करें..."
                  />
                </Field>

                {vendorAbstract.extracted && Object.keys(vendorAbstract.extracted).length > 0 && (
                  <div className="panel">
                    <h3>पहचानी गई Vendor Bill Details</h3>
                    <p>Previous Amount: {money(vendorAbstract.extracted.previous_amount)}</p>
                    <p>Current Amount: {money(vendorAbstract.extracted.current_amount)}</p>
                    <p>Cumulative Amount: {money(vendorAbstract.extracted.cumulative_amount)}</p>
                    <p>GST: {money(vendorAbstract.extracted.gst_amount)}</p>
                    <p>Gross Bill: {money(vendorAbstract.extracted.gross_bill)}</p>
                    <p>NET Payable: {money(vendorAbstract.extracted.net_payable)}</p>
                  </div>
                )}
              </div>

              <div className="panel">
                <h2>Vendor Bill Details</h2>

                <form className="entry-grid" onSubmit={saveVendorBill}>
                  <Field label="Agency Name">
                    <input
                      required
                      value={vendorForm.agency_name}
                      onChange={e => setVendorForm({
                        ...vendorForm,
                        agency_name: e.target.value
                      })}
                      placeholder="P E ERECTORS PRIVATE LIMITED"
                    />
                  </Field>

                  <Field label="Bill Number">
                    <input
                      required
                      value={vendorForm.bill_no}
                      onChange={e => setVendorForm({
                        ...vendorForm,
                        bill_no: e.target.value
                      })}
                      placeholder="RA-17"
                    />
                  </Field>

                  <Field label="Bill Date">
                    <input
                      type="date"
                      value={vendorForm.bill_date}
                      onChange={e => setVendorForm({
                        ...vendorForm,
                        bill_date: e.target.value
                      })}
                    />
                  </Field>

                  <Field label="Billing Period">
                    <input
                      value={vendorForm.billing_period}
                      onChange={e => setVendorForm({
                        ...vendorForm,
                        billing_period: e.target.value
                      })}
                      placeholder="26.09.2026 to 25.10.2026"
                    />
                  </Field>

                  <Field label="Grand Total Amount">
                    <input
                      type="number"
                      step="any"
                      value={vendorForm.total_amount}
                      onChange={e => setVendorForm({
                        ...vendorForm,
                        total_amount: e.target.value
                      })}
                    />
                  </Field>

                  <Field label="Net Payable Amount">
                    <input
                      type="number"
                      step="any"
                      value={vendorForm.net_payable}
                      onChange={e => setVendorForm({
                        ...vendorForm,
                        net_payable: e.target.value
                      })}
                    />
                  </Field>

                  <div className="form-end">
                    <button className="primary" disabled={busy}>
                      <Plus size={16} />
                      Save / Update Vendor Bill
                    </button>
                  </div>
                </form>
              </div>

              <div className="panel">
                <h2>Vendor Bill Register</h2>
                <BillTable
                  bills={bills}
                  type="vendor"
                  paidFor={paidFor}
                  outstanding={vendorOutstanding}
                  onView={setSelectedBill}
                />
              </div>

              <div className="panel">
                <h2>Vendor Payment Entry</h2>

                <form className="entry-grid" onSubmit={addVendorPayment}>
                  <Field label="Select Vendor Bill">
                    <select
                      required
                      value={paymentForm.bill_id}
                      onChange={e => setPaymentForm({
                        ...paymentForm,
                        bill_id: e.target.value
                      })}
                    >
                      <option value="">Select Bill</option>
                      {bills.map(b => (
                        <option key={b.id} value={b.id}>
                          {b.agency_name} - {b.bill_no}
                        </option>
                      ))}
                    </select>
                  </Field>

                  <Field label="Payment Date">
                    <input
                      type="date"
                      required
                      value={paymentForm.payment_date}
                      onChange={e => setPaymentForm({
                        ...paymentForm,
                        payment_date: e.target.value
                      })}
                    />
                  </Field>

                  <Field label="Payment Amount">
                    <input
                      type="number"
                      min="0.01"
                      step="any"
                      required
                      value={paymentForm.amount}
                      onChange={e => setPaymentForm({
                        ...paymentForm,
                        amount: e.target.value
                      })}
                    />
                  </Field>

                  <Field label="UTR / Reference No.">
                    <input
                      value={paymentForm.reference_no}
                      onChange={e => setPaymentForm({
                        ...paymentForm,
                        reference_no: e.target.value
                      })}
                    />
                  </Field>

                  <Field label="Remarks">
                    <input
                      value={paymentForm.remarks}
                      onChange={e => setPaymentForm({
                        ...paymentForm,
                        remarks: e.target.value
                      })}
                    />
                  </Field>

                  <div className="form-end">
                    <button className="primary">
                      <IndianRupee size={16} />
                      Save Payment
                    </button>
                  </div>
                </form>

                <DataTable
                  rows={vendorPayments}
                  columns={[
                    'payment_date', 'bill_id', 'amount',
                    'reference_no', 'remarks'
                  ]}
                />
              </div>
            </>
          )}

          {tab === 'customer' && !customerUnlocked && (
            <div className="panel" style={{maxWidth:520, margin:'50px auto', textAlign:'center'}}>
              <LockKeyhole size={38}/>
              <h1>Customer Bill Locked</h1>
              <p>Customer Bill सेक्शन खोलने के लिए पासवर्ड डालें।</p>
              <form onSubmit={e => { e.preventDefault(); if (customerLockInput === '22334455') { setCustomerUnlocked(true); setCustomerLockInput(''); setNotice('Customer Bill खुल गया।'); } else setNotice('गलत पासवर्ड। कृपया दोबारा प्रयास करें।'); }}>
                <input type="password" value={customerLockInput} onChange={e => setCustomerLockInput(e.target.value)} placeholder="Customer Bill Password" autoComplete="off" required/>
                <button className="primary full" type="submit"><LockKeyhole size={16}/> Unlock Customer Bill</button>
              </form>
              {notice && <p>{notice}</p>}
            </div>
          )}

          {tab === 'customer' && customerUnlocked && (
            <>
              <div className="welcome">
                <div>
                  <span className="eyebrow">CUSTOMER BILLING</span>
                  <h1>Customer Bill Management</h1>
                  <p>Import BHEL Customer RA Abstract Excel or paste its rows.</p>
                </div>

                <button
                  className="secondary"
                  onClick={() => exportRows(customerBills.map(b => ({
                    project_name: b.project_name,
                    contract_ref: b.contract_ref,
                    bill_no: b.bill_no,
                    bill_date: b.bill_date,
                    billing_period: b.billing_period,
                    total_amount: b.total_amount
                  })), 'Customer_Bills')}
                >
                  <Download size={16} />
                  Export Register
                </button>
              </div>

              <div className="stats">
                <Stat title="Total Bills" value={customerMonthBills.length} />
                <Stat title="Total Billed" value={money(customerTotals.billed)} />
                <Stat title="Received" value={money(customerTotals.received)} />
                <Stat title="Outstanding" value={money(customerTotals.billed - customerTotals.received)} />
              </div>

              <div className="panel">
                <h2>Import Customer Abstract</h2>
                <p>अपनी मौजूदा BHEL Customer RA Bill Excel फाइल चुनें।</p>

                <label className="secondary upload">
                  <Upload size={16} />
                  Import Customer Excel
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={e => readBillFile(e, 'customer')}
                  />
                </label>

                {customerOriginalFile && <p className="success-text">Original file: {customerOriginalFile.name} · {Math.ceil(customerOriginalFile.size / 1024)} KB</p>}

                {customerAbstract.rows.length > 0 && (
                  <p className="success-text">
                    Excel तैयार: {customerAbstract.sheetName} · {customerAbstract.rows.length} rows
                  </p>
                )}

                <Field label="या Excel से पूरा Abstract Copy करके यहाँ Paste करें">
                  <textarea
                    rows="5"
                    value={customerPaste}
                    onChange={e => handleBillPaste(e.target.value, 'customer')}
                    placeholder="Excel की पूरी टेबल Copy करके यहाँ Paste करें..."
                  />
                </Field>

                {customerAbstract.extracted && Object.keys(customerAbstract.extracted).length > 0 && (
                  <div className="panel">
                    <h3>पहचानी गई Customer Bill Details</h3>
                    <p>Previous Amount: {money(customerAbstract.extracted.previous_amount)}</p>
                    <p>Current Amount: {money(customerAbstract.extracted.current_amount)}</p>
                    <p>Cumulative Amount: {money(customerAbstract.extracted.cumulative_amount)}</p>
                    <p>GST: {money(customerAbstract.extracted.gst_amount)}</p>
                    <p>Gross Bill: {money(customerAbstract.extracted.gross_bill)}</p>
                  </div>
                )}
              </div>

              <div className="panel">
                <h2>Customer Bill Details</h2>

                <form className="entry-grid" onSubmit={saveCustomerBill}>
                  <Field label="Project Name">
                    <input
                      value={customerForm.project_name}
                      onChange={e => setCustomerForm({
                        ...customerForm,
                        project_name: e.target.value
                      })}
                    />
                  </Field>

                  <Field label="Contract Reference">
                    <input
                      required
                      value={customerForm.contract_ref}
                      onChange={e => setCustomerForm({
                        ...customerForm,
                        contract_ref: e.target.value
                      })}
                    />
                  </Field>

                  <Field label="Bill Number">
                    <input
                      required
                      value={customerForm.bill_no}
                      onChange={e => setCustomerForm({
                        ...customerForm,
                        bill_no: e.target.value
                      })}
                      placeholder="RA-16"
                    />
                  </Field>

                  <Field label="Bill Date">
                    <input
                      type="date"
                      value={customerForm.bill_date}
                      onChange={e => setCustomerForm({
                        ...customerForm,
                        bill_date: e.target.value
                      })}
                    />
                  </Field>

                  <Field label="Billing Period">
                    <input
                      value={customerForm.billing_period}
                      onChange={e => setCustomerForm({
                        ...customerForm,
                        billing_period: e.target.value
                      })}
                      placeholder="01/09/2026 To 30/09/2026"
                    />
                  </Field>

                  <Field label="Total Amount Claimed">
                    <input
                      type="number"
                      step="any"
                      value={customerForm.total_amount}
                      onChange={e => setCustomerForm({
                        ...customerForm,
                        total_amount: e.target.value
                      })}
                    />
                  </Field>

                  <div className="form-end">
                    <button className="primary" disabled={busy}>
                      <Plus size={16} />
                      Save / Update Customer Bill
                    </button>
                  </div>
                </form>
              </div>

              <div className="panel">
                <h2>Customer Bill Register</h2>
                <BillTable
                  bills={customerBills}
                  type="customer"
                  paidFor={receivedFor}
                  outstanding={customerOutstanding}
                  onView={setSelectedBill}
                />
              </div>

              <div className="panel">
                <h2>Customer Receipt Entry</h2>

                <form className="entry-grid" onSubmit={addCustomerReceipt}>
                  <Field label="Select Customer Bill">
                    <select
                      required
                      value={receiptForm.bill_id}
                      onChange={e => setReceiptForm({
                        ...receiptForm,
                        bill_id: e.target.value
                      })}
                    >
                      <option value="">Select Bill</option>
                      {customerBills.map(b => (
                        <option key={b.id} value={b.id}>
                          {b.bill_no} - {b.billing_period}
                        </option>
                      ))}
                    </select>
                  </Field>

                  <Field label="Receipt Date">
                    <input
                      type="date"
                      required
                      value={receiptForm.receipt_date}
                      onChange={e => setReceiptForm({
                        ...receiptForm,
                        receipt_date: e.target.value
                      })}
                    />
                  </Field>

                  <Field label="Received Amount">
                    <input
                      type="number"
                      min="0.01"
                      step="any"
                      required
                      value={receiptForm.amount}
                      onChange={e => setReceiptForm({
                        ...receiptForm,
                        amount: e.target.value
                      })}
                    />
                  </Field>

                  <Field label="Bank / Reference No.">
                    <input
                      value={receiptForm.reference_no}
                      onChange={e => setReceiptForm({
                        ...receiptForm,
                        reference_no: e.target.value
                      })}
                    />
                  </Field>

                  <Field label="Remarks">
                    <input
                      value={receiptForm.remarks}
                      onChange={e => setReceiptForm({
                        ...receiptForm,
                        remarks: e.target.value
                      })}
                    />
                  </Field>

                  <div className="form-end">
                    <button className="primary">
                      <IndianRupee size={16} />
                      Save Receipt
                    </button>
                  </div>
                </form>

                <DataTable
                  rows={customerReceipts}
                  columns={[
                    'receipt_date', 'bill_id', 'amount',
                    'reference_no', 'remarks'
                  ]}
                />
              </div>
            </>
          )}

          {tab === 'drawings' && (
            <>
              <div className="welcome">
                <div>
                  <span className="eyebrow">DOCUMENT LIBRARY</span>
                  <h1>Drawing Folder</h1>
                  <p>Store and organize drawings by department.</p>
                </div>

                <label className="primary upload">
                  <Upload size={16} />
                  Upload Drawing
                  <input
                    type="file"
                    accept=".pdf,.dwg,.dxf,.png,.jpg,.jpeg,.xlsx"
                    onChange={uploadDrawing}
                  />
                </label>
              </div>

              <div className="panel">
                <Field label="Drawing Department">
                  <select
                    value={form.department}
                    onChange={e => setForm({
                      ...form,
                      department: e.target.value
                    })}
                  >
                    {DEPTS.map(d => <option key={d}>{d}</option>)}
                  </select>
                </Field>

                <div className="dept-grid">
                  {DEPTS.map(d => (
                    <div className="dept" key={d}>
                      <b>{d}</b>
                      <small>
                        {docs.filter(x => x.department === d).length} files
                      </small>
                    </div>
                  ))}
                </div>
              </div>

              <div className="panel">
                <h2>Saved Drawings</h2>

                {docs.map(f => (
                  <div className="file-row" key={f.id}>
                    <FolderOpen />

                    <div>
                      <b>{f.file_name}</b>
                      <small>{f.department}</small>
                    </div>

                    <button
                      className="link"
                      onClick={async () => {
                        const { data, error } = await supabase.storage
                          .from('drawings')
                          .createSignedUrl(f.storage_path, 120);

                        if (error) {
                          setNotice(error.message);
                        } else {
                          window.open(data.signedUrl, '_blank');
                        }
                      }}
                    >
                      Open
                    </button>
                  </div>
                ))}

                {!docs.length && (
                  <div className="empty">No drawings uploaded yet.</div>
                )}
              </div>
            </>
          )}

          {selectedBill && (
            <div className="panel bill-view">
              <div className="panel-head">
                <h2>
                  {selectedBill.type === 'vendor' ? 'Vendor' : 'Customer'} Bill Abstract
                </h2>

                <div className="actions">
                  <button className="secondary" onClick={printAbstract}>
                    <Printer size={16} />
                    Print / PDF
                  </button>

                  <button
                    className="secondary"
                    onClick={() => setSelectedBill(null)}
                  >
                    <X size={16} />
                    Close
                  </button>
                </div>
              </div>

              <p><b>Bill No:</b> {normalizeRaBillNo(selectedBill.bill.bill_no)}</p>

              <p>
                <b>Bill Date:</b> {selectedBill.bill.bill_date}
                {'　'}
                <b>Billing Period:</b> {selectedBill.bill.billing_period}
              </p>

              {selectedBill.type === 'vendor'
                ? (
                  <p>
                    <b>Agency:</b> {selectedBill.bill.agency_name}
                  </p>
                )
                : (
                  <p>
                    <b>Project:</b> {selectedBill.bill.project_name}
                    {'　'}
                    <b>Contract:</b> {selectedBill.bill.contract_ref}
                  </p>
                )}

              <p>
                <b>Total Amount:</b> {money(selectedBill.bill.total_amount)}
              </p>

              {selectedBill.type === 'vendor' && (
                <p>
                  <b>Net Payable:</b> {money(selectedBill.bill.net_payable)}
                </p>
              )}

              {selectedBill.bill.abstract_data?.archive?.storage_path && <p><button className="secondary" onClick={async () => { const { data, error } = await supabase.storage.from('drawings').createSignedUrl(selectedBill.bill.abstract_data.archive.storage_path, 120); if (error) setNotice(error.message); else window.open(data.signedUrl, '_blank'); }}><Download size={15}/> Original Excel खोलें ({selectedBill.bill.abstract_data.archive.file_name})</button></p>}
              <div className="table-wrap">
                <table style={{minWidth: selectedBill.type === 'vendor' ? 1700 : 1250}}>
                  <tbody>
                    {billRows(selectedBill.bill).map((r, i) => (
                      <tr key={i}>
                        {(Array.isArray(r) ? r : Object.values(r || {})).map((v, j) => {
                          const merges = selectedBill.bill.abstract_data?.merges || [];
                          const covered = merges.some(m =>
                            i >= m.s.r && i <= m.e.r && j >= m.s.c && j <= m.e.c &&
                            !(i === m.s.r && j === m.s.c)
                          );
                          if (covered) return null;
                          const merge = merges.find(m => m.s.r === i && m.s.c === j);
                          const colSpan = merge ? merge.e.c - merge.s.c + 1 : 1;
                          const rowSpan = merge ? merge.e.r - merge.s.r + 1 : 1;
                          return <td key={j} colSpan={colSpan} rowSpan={rowSpan}>{String(v ?? '')}</td>;
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      </main>
    </div>
  );
}

function Setup() {
  return (
    <div className="login-wrap">
      <div className="login-card">
        <Factory size={38} />
        <h1>Boiler Management</h1>
        <p>Supabase connection required</p>
        <div className="setup-error">
          Project URL या Publishable Key सेट नहीं है।
        </div>
      </div>
    </div>
  );
}

function Stat({ title, value }) {
  return (
    <div className="stat">
      <small>{title}</small>
      <strong>{value}</strong>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="field">
      <span>{label}</span>
      {children}
    </label>
  );
}

function DprTable({ rows, onDelete }) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {[
              'Date', 'Dept.', 'PG', 'MA', 'DU', 'Description',
              'Qty', 'Unit Weight', 'Total Weight', 'Pre-Assembly',
              'Placement', 'Alignment', 'Bolting', 'Remarks', ''
            ].map(x => <th key={x}>{x}</th>)}
          </tr>
        </thead>

        <tbody>
          {rows.map(r => (
            <tr key={r.id}>
              <td>{r.work_date}</td>
              <td>{r.department}</td>
              <td>{r.pg}</td>
              <td>{r.ma}</td>
              <td>{r.du}</td>
              <td>{r.description}</td>
              <td>{r.quantity}</td>
              <td>{r.unit_weight ?? 0}</td>
              <td>{weight(r.quantity, r.unit_weight)}</td>
              <td>{r.pre_assembly}</td>
              <td>{r.placement}</td>
              <td>{r.alignment}</td>
              <td>{r.bolting}</td>
              <td>{r.remark}</td>
              <td>
                <button
                  className="delete"
                  onClick={() => onDelete(r.id)}
                >
                  <Trash2 size={15} />
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      {!rows.length && (
        <div className="empty">No records found.</div>
      )}
    </div>
  );
}

function DataTable({ rows, columns }) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {columns.map(c => (
              <th key={c}>{c.replaceAll('_', ' ')}</th>
            ))}
          </tr>
        </thead>

        <tbody>
          {rows.map((r, i) => (
            <tr key={r.id || i}>
              {columns.map(c => (
                <td key={c}>{r[c] ?? ''}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>

      {!rows.length && (
        <div className="empty">No records yet.</div>
      )}
    </div>
  );
}

function BillTable({ bills, type, paidFor, outstanding, onView }) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Bill No.</th>
            <th>{type === 'vendor' ? 'Agency' : 'Project'}</th>
            <th>Bill Date</th>
            <th>Billing Period</th>
            <th>Bill Amount</th>
            <th>{type === 'vendor' ? 'Paid' : 'Received'}</th>
            <th>Outstanding</th>
            <th>View</th>
          </tr>
        </thead>

        <tbody>
          {bills.map(b => (
            <tr key={b.id}>
              <td><button className="link" onClick={() => onView({ bill: b, type })}>{normalizeRaBillNo(b.bill_no)}</button></td>
              <td>
                {type === 'vendor' ? b.agency_name : b.project_name}
              </td>
              <td>{b.bill_date}</td>
              <td>{b.billing_period}</td>
              <td>
                {money(
                  type === 'vendor'
                    ? (b.net_payable || b.total_amount)
                    : b.total_amount
                )}
              </td>
              <td>{money(paidFor(b.id))}</td>
              <td>{money(outstanding(b))}</td>
              <td>
                <button
                  className="secondary"
                  onClick={() => onView({ bill: b, type })}
                >
                  <Eye size={15} />
                  View
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      {!bills.length && (
        <div className="empty">No bills saved yet.</div>
      )}
    </div>
  );
}

export default App;
