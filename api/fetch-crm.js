// ============================================================
//  fetch-crm.js  —  Vercel serverless function
//  Receives CRM type + credentials, fetches data, returns Excel
//  Supports: erpnext, odoo, hubspot, zoho, pipedrive, freshsales, salesforce
// ============================================================

const ExcelJS = require('exceljs');
const axios   = require('axios');
const qs      = require('qs');
const { reportError } = require('./report-error');

// ── PALETTE (matches Stock Report Tool exactly) ───────────────
const C = {
  darkBlue:    '1F3864',
  midBlue:     '2F5496',
  lightBlue:   'BDD7EE',
  headerFont:  'FFFFFF',
  altRow:      'EBF3FB',
  white:       'FFFFFF',
  orange:      'ED7D31',
  green:       '70AD47',
  red:         'C00000',
  yellow:      'FFF2CC',
  borderColor: '9DC3E6',
  summaryBg:   'DEEAF1',
  purple:      '7030A0',
  teal:        '00B0A0',
};

// ── BRAND COLORS (cycles through 20 colours — matches office Stock Report Tool) ──
const BRAND_COLORS = [
  'FF4472C4','FF70AD47','FFFFC000','FFED7D31','FF5B9BD5',
  'FF9DC3E6','FFA9D18E','FFD9D9D9','FF843C0C','FF375623',
  'FF002060','FF2E75B6','FF7030A0','FFBF8F00','FFBF3F00',
  'FF215868','FF538135','FFBF4040','FF0070C0','FF404040',
];

// ── SAFE FETCH WRAPPER ────────────────────────────────────────
// Wraps any axios call — on failure returns empty data + pushes a warning
async function safeGet(label, axiosPromise, fallback, warnings) {
  try {
    const res = await axiosPromise;
    return res;
  } catch (err) {
    const status = err.response ? err.response.status : null;
    let reason = err.message || 'Unknown error';
    if (status === 401) reason = `${label}: Token invalid or expired (401)`;
    else if (status === 403) reason = `${label}: Missing permissions/scopes (403)`;
    else if (status === 404) reason = `${label}: Endpoint not found — check your URL (404)`;
    else if (status === 429) reason = `${label}: Rate limit exceeded (429) — try again in 60s`;
    else if (status === 500) reason = `${label}: CRM server error (500)`;
    else if (err.code === 'ECONNABORTED') reason = `${label}: Request timed out`;
    else reason = `${label}: ${reason}`;
    warnings.push(reason);
    return fallback;
  }
}

// ── STYLE HELPERS ─────────────────────────────────────────────
function thinBorder() {
  const side = { style: 'thin', color: { argb: 'FF' + C.borderColor } };
  return { top: side, left: side, bottom: side, right: side };
}
function headerStyle(bgHex, fgHex = C.headerFont) {
  return {
    font:      { bold: true, color: { argb: 'FF' + fgHex }, size: 11, name: 'Calibri' },
    fill:      { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + bgHex } },
    alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
    border:    thinBorder(),
  };
}
function cellStyle(bgHex, bold = false, align = 'left') {
  return {
    font:      { bold, size: 10, name: 'Calibri', color: { argb: 'FF222222' } },
    fill:      { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + bgHex } },
    alignment: { horizontal: align, vertical: 'middle' },
    border:    thinBorder(),
  };
}
// DD MM YYYY — Arun's standard for every date in every sheet
function toStdDate(val) {
  if (!val) return '';
  const d = (val instanceof Date) ? val : new Date(val);
  if (isNaN(d)) return String(val);
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${dd} ${mm} ${d.getFullYear()}`;
}
function fmtDate() {
  return toStdDate(new Date());
}
// MM YYYY display for month aggregates — e.g. "2026-01" → "01 2026"
function toMonthDisplay(monthKey) {
  if (!monthKey || monthKey.length < 7) return monthKey || '';
  const mm = monthKey.substring(5, 7);
  const yyyy = monthKey.substring(0, 4);
  return `${mm} ${yyyy}`;
}
function setColWidths(sheet, widths) {
  widths.forEach((w, i) => { sheet.getColumn(i + 1).width = w; });
}
function freezeRow(sheet, row) {
  sheet.views = [{ state: 'frozen', ySplit: row }];
}
function addSheetHeader(ws, title, subtitle, cols) {
  ws.views = [{ showGridLines: false }];
  const span = `A1:${String.fromCharCode(64 + cols)}1`;
  ws.mergeCells(span);
  const t = ws.getCell('A1');
  t.value = title;
  t.style = { font: { bold: true, size: 16, name: 'Calibri', color: { argb: 'FF' + C.headerFont } },
               fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.darkBlue } },
               alignment: { horizontal: 'center', vertical: 'middle' } };
  ws.getRow(1).height = 36;

  const span2 = `A2:${String.fromCharCode(64 + cols)}2`;
  ws.mergeCells(span2);
  const s = ws.getCell('A2');
  s.value = `${subtitle}  |  Generated: ${fmtDate()}  |  Powered by CRM Data Extractor`;
  s.style = { font: { bold: false, size: 10, name: 'Calibri', color: { argb: 'FF' + C.headerFont } },
               fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.midBlue } },
               alignment: { horizontal: 'center', vertical: 'middle' } };
  ws.getRow(2).height = 20;

  ws.getRow(3).height = 8; // spacer
}

// ── CRM FETCHERS ──────────────────────────────────────────────

async function fetchHubSpot(apiKey, reportType) {
  const headers = { Authorization: `Bearer ${apiKey}` };
  const base    = 'https://api.hubapi.com';
  const rt      = (reportType || 'full').toLowerCase();
  const warnings = [];

  const needDeals      = ['full','pipeline'].includes(rt);
  const needContacts   = ['full','contacts'].includes(rt);
  const needActivities = ['full','activities'].includes(rt);

  const emptyResults = { data: { results: [] } };
  const [dealsRes, contactsRes, activitiesRes] = await Promise.all([
    needDeals      ? safeGet('Deals',      axios.get(`${base}/crm/v3/objects/deals?limit=100&properties=dealname,amount,dealstage,closedate,pipeline,hs_deal_stage_probability,createdate`, { headers }), emptyResults, warnings) : Promise.resolve(emptyResults),
    needContacts   ? safeGet('Contacts',   axios.get(`${base}/crm/v3/objects/contacts?limit=100&properties=firstname,lastname,email,phone,company,jobtitle,lastmodifieddate,hs_lead_status`, { headers }), emptyResults, warnings) : Promise.resolve(emptyResults),
    needActivities ? safeGet('Activities', axios.get(`${base}/crm/v3/objects/tasks?limit=100&properties=hs_task_subject,hs_task_status,hs_timestamp,hs_task_type`, { headers }), emptyResults, warnings) : Promise.resolve(emptyResults),
  ]);

  const deals    = dealsRes.data.results     || [];
  const contacts = contactsRes.data.results  || [];
  const tasks    = activitiesRes.data.results || [];

  const pipeline = deals.map(d => ({
    name: d.properties.dealname||'', amount: parseFloat(d.properties.amount)||0,
    stage: d.properties.dealstage||'', closeDate: d.properties.closedate||'',
    probability: parseFloat(d.properties.hs_deal_stage_probability)||0, pipeline: d.properties.pipeline||'',
  }));
  const contactList = contacts.map(c => ({
    name: `${c.properties.firstname||''} ${c.properties.lastname||''}`.trim(),
    email: c.properties.email||'', phone: c.properties.phone||'',
    company: c.properties.company||'', title: c.properties.jobtitle||'',
    lastActivity: c.properties.lastmodifieddate||'', status: c.properties.hs_lead_status||'Active',
  }));
  const activities = tasks.map(t => ({
    subject: t.properties.hs_task_subject||'', status: t.properties.hs_task_status||'',
    dueDate: t.properties.hs_timestamp||'', type: t.properties.hs_task_type||'',
  }));
  const wonDeals  = deals.filter(d=>d.properties.dealstage==='closedwon').map(d=>({ name:d.properties.dealname, amount:parseFloat(d.properties.amount)||0, result:'Won', closeDate:d.properties.closedate }));
  const lostDeals = deals.filter(d=>d.properties.dealstage==='closedlost').map(d=>({ name:d.properties.dealname, amount:parseFloat(d.properties.amount)||0, result:'Lost', closeDate:d.properties.closedate }));

  const totalPipeline = pipeline.reduce((s,d)=>s+d.amount,0);
  const openDeals     = pipeline.filter(d=>!['closedwon','closedlost'].includes(d.stage)).length;
  const closingSoon   = pipeline.filter(d=>{ if(!d.closeDate)return false; const days=(new Date(d.closeDate)-new Date())/86400000; return days>=0&&days<=30; }).length;
  const wonThisMonth  = wonDeals.filter(d=>{ if(!d.closeDate)return false; const cd=new Date(d.closeDate); const now=new Date(); return cd.getMonth()===now.getMonth()&&cd.getFullYear()===now.getFullYear(); }).length;

  return {
    crmName: 'HubSpot', reportType: rt, warnings,
    kpis: { totalPipeline, openDeals, closingSoon, wonThisMonth, totalContacts: contactList.length, totalActivities: tasks.length },
    pipeline, leads: contactList.filter(c=>c.status&&c.status.toLowerCase().includes('lead')),
    contacts: contactList, activities, wonLost: [...wonDeals, ...lostDeals],
  };
}

async function fetchZoho(apiKey, reportType) {
  const headers  = { 'Authorization': `Zoho-oauthtoken ${apiKey}` };
  const base     = 'https://www.zohoapis.com/crm/v2';
  const rt       = (reportType || 'full').toLowerCase();
  const warnings = [];

  const needDeals    = ['full','pipeline'].includes(rt);
  const needLeads    = ['full','leads'].includes(rt);
  const needContacts = ['full','contacts'].includes(rt);

  const emptyData = { data: { data: [] } };
  const [dealsRes, leadsRes, contactsRes] = await Promise.all([
    needDeals    ? safeGet('Deals',    axios.get(`${base}/Deals?fields=Deal_Name,Amount,Stage,Closing_Date,Probability,Account_Name`, { headers }), emptyData, warnings) : Promise.resolve(emptyData),
    needLeads    ? safeGet('Leads',    axios.get(`${base}/Leads?fields=First_Name,Last_Name,Email,Phone,Company,Lead_Status,Last_Activity_Time`, { headers }), emptyData, warnings) : Promise.resolve(emptyData),
    needContacts ? safeGet('Contacts', axios.get(`${base}/Contacts?fields=First_Name,Last_Name,Email,Phone,Account_Name,Title,Last_Activity_Time`, { headers }), emptyData, warnings) : Promise.resolve(emptyData),
  ]);

  const deals    = dealsRes.data.data    || [];
  const leads    = leadsRes.data.data    || [];
  const contacts = contactsRes.data.data || [];

  const pipeline    = deals.map(d => ({ name:d.Deal_Name||'', amount:parseFloat(d.Amount)||0, stage:d.Stage||'', closeDate:d.Closing_Date||'', probability:parseFloat(d.Probability)||0, pipeline:d.Account_Name||'' }));
  const leadList    = leads.map(l => ({ name:`${l.First_Name||''} ${l.Last_Name||''}`.trim(), email:l.Email||'', phone:l.Phone||'', company:l.Company||'', status:l.Lead_Status||'', lastActivity:l.Last_Activity_Time||'' }));
  const contactList = contacts.map(c => ({ name:`${c.First_Name||''} ${c.Last_Name||''}`.trim(), email:c.Email||'', phone:c.Phone||'', company:c.Account_Name||'', title:c.Title||'', lastActivity:c.Last_Activity_Time||'', status:'Active' }));
  const wonDeals    = deals.filter(d=>d.Stage==='Closed Won').map(d=>({ name:d.Deal_Name, amount:parseFloat(d.Amount)||0, result:'Won', closeDate:d.Closing_Date }));
  const lostDeals   = deals.filter(d=>d.Stage==='Closed Lost').map(d=>({ name:d.Deal_Name, amount:parseFloat(d.Amount)||0, result:'Lost', closeDate:d.Closing_Date }));

  const totalPipeline = pipeline.reduce((s,d)=>s+d.amount,0);
  const openDeals     = pipeline.filter(d=>!['Closed Won','Closed Lost'].includes(d.stage)).length;
  const closingSoon   = pipeline.filter(d=>{ if(!d.closeDate)return false; const days=(new Date(d.closeDate)-new Date())/86400000; return days>=0&&days<=30; }).length;
  const wonThisMonth  = wonDeals.filter(d=>{ if(!d.closeDate)return false; const cd=new Date(d.closeDate); const now=new Date(); return cd.getMonth()===now.getMonth()&&cd.getFullYear()===now.getFullYear(); }).length;

  return { crmName:'Zoho CRM', reportType:rt, warnings, kpis:{totalPipeline,openDeals,closingSoon,wonThisMonth,totalLeads:leads.length,totalContacts:contacts.length}, pipeline, leads:leadList, contacts:contactList, activities:[], wonLost:[...wonDeals,...lostDeals] };
}

async function fetchPipedrive(apiKey, reportType) {
  const base     = 'https://api.pipedrive.com/v1';
  const p        = `api_token=${apiKey}`;
  const rt       = (reportType || 'full').toLowerCase();
  const warnings = [];

  const needDeals    = ['full','pipeline'].includes(rt);
  const needContacts = ['full','contacts'].includes(rt);

  const emptyData = { data: { data: [] } };
  const [dealsRes, personsRes] = await Promise.all([
    needDeals    ? safeGet('Deals',   axios.get(`${base}/deals?${p}&limit=100&status=all_not_deleted`), emptyData, warnings) : Promise.resolve(emptyData),
    needContacts ? safeGet('Persons', axios.get(`${base}/persons?${p}&limit=100`), emptyData, warnings) : Promise.resolve(emptyData),
  ]);

  const deals   = dealsRes.data.data   || [];
  const persons = personsRes.data.data || [];

  const pipeline    = deals.map(d => ({ name:d.title||'', amount:parseFloat(d.value)||0, stage:d.stage_name||'', closeDate:d.close_time||d.expected_close_date||'', probability:parseFloat(d.probability)||0, pipeline:d.pipeline_name||'' }));
  const contactList = persons.map(p => ({ name:p.name||'', email:(p.email&&p.email[0]&&p.email[0].value)||'', phone:(p.phone&&p.phone[0]&&p.phone[0].value)||'', company:p.org_name||'', title:p.job_title||'', lastActivity:p.last_activity_date||'', status:'Active' }));
  const wonDeals    = deals.filter(d=>d.status==='won').map(d=>({ name:d.title, amount:parseFloat(d.value)||0, result:'Won', closeDate:d.close_time }));
  const lostDeals   = deals.filter(d=>d.status==='lost').map(d=>({ name:d.title, amount:parseFloat(d.value)||0, result:'Lost', closeDate:d.close_time }));

  const totalPipeline = pipeline.reduce((s,d)=>s+d.amount,0);
  const openDeals     = deals.filter(d=>d.status==='open').length;
  const closingSoon   = pipeline.filter(d=>{ if(!d.closeDate)return false; const days=(new Date(d.closeDate)-new Date())/86400000; return days>=0&&days<=30; }).length;
  const wonThisMonth  = wonDeals.filter(d=>{ if(!d.closeDate)return false; const cd=new Date(d.closeDate); const now=new Date(); return cd.getMonth()===now.getMonth()&&cd.getFullYear()===now.getFullYear(); }).length;

  return { crmName:'Pipedrive', reportType:rt, warnings, kpis:{totalPipeline,openDeals,closingSoon,wonThisMonth,totalContacts:persons.length}, pipeline, leads:[], contacts:contactList, activities:[], wonLost:[...wonDeals,...lostDeals] };
}

async function fetchFreshsales(apiKey, domain, reportType) {
  const base     = `https://${domain}.myfreshworks.com/crm/sales/api`;
  const headers  = { Authorization: `Token token=${apiKey}` };
  const rt       = (reportType || 'full').toLowerCase();
  const warnings = [];

  const needDeals    = ['full','pipeline'].includes(rt);
  const needContacts = ['full'].includes(rt);
  const needLeads    = ['full','leads'].includes(rt);

  const [dealsRes, contactsRes, leadsRes] = await Promise.all([
    needDeals    ? safeGet('Deals',    axios.get(`${base}/deals?include=deal_stage,deal_type`, { headers }), { data:{deals:[]} },    warnings) : Promise.resolve({ data:{deals:[]} }),
    needContacts ? safeGet('Contacts', axios.get(`${base}/contacts?include=owner`, { headers }),              { data:{contacts:[]} }, warnings) : Promise.resolve({ data:{contacts:[]} }),
    needLeads    ? safeGet('Leads',    axios.get(`${base}/leads?include=owner`, { headers }),                 { data:{leads:[]} },    warnings) : Promise.resolve({ data:{leads:[]} }),
  ]);

  const deals    = dealsRes.data.deals       || [];
  const contacts = contactsRes.data.contacts || [];
  const leads    = leadsRes.data.leads       || [];

  const pipeline    = deals.map(d => ({ name:d.name||'', amount:parseFloat(d.amount)||0, stage:d.deal_stage?d.deal_stage.name:'', closeDate:d.expected_close||'', probability:parseFloat(d.probability)||0, pipeline:'' }));
  const contactList = contacts.map(c => ({ name:c.display_name||'', email:c.email||'', phone:c.mobile_number||c.work_number||'', company:c.company_name||'', title:c.job_title||'', lastActivity:c.last_seen||'', status:'Active' }));
  const leadList    = leads.map(l => ({ name:l.display_name||'', email:l.email||'', phone:l.mobile_number||'', company:l.company||'', status:l.lead_stage_name||'', lastActivity:l.last_seen||'' }));
  const wonDeals    = deals.filter(d=>d.deal_stage&&d.deal_stage.name==='Won').map(d=>({ name:d.name, amount:parseFloat(d.amount)||0, result:'Won', closeDate:d.expected_close }));
  const lostDeals   = deals.filter(d=>d.deal_stage&&d.deal_stage.name==='Lost').map(d=>({ name:d.name, amount:parseFloat(d.amount)||0, result:'Lost', closeDate:d.expected_close }));

  const totalPipeline = pipeline.reduce((s,d)=>s+d.amount,0);
  const openDeals     = pipeline.filter(d=>!['Won','Lost'].includes(d.stage)).length;
  const closingSoon   = pipeline.filter(d=>{ if(!d.closeDate)return false; const days=(new Date(d.closeDate)-new Date())/86400000; return days>=0&&days<=30; }).length;
  const wonThisMonth  = wonDeals.filter(d=>{ if(!d.closeDate)return false; const cd=new Date(d.closeDate); const now=new Date(); return cd.getMonth()===now.getMonth()&&cd.getFullYear()===now.getFullYear(); }).length;

  return { crmName:'Freshsales', reportType:rt, warnings, kpis:{totalPipeline,openDeals,closingSoon,wonThisMonth,totalLeads:leads.length}, pipeline, leads:leadList, contacts:contactList, activities:[], wonLost:[...wonDeals,...lostDeals] };
}

async function fetchERPNext(url, username, password, reportType, fiscalYear) {
  const base     = url.replace(/\/$/, '');
  const warnings = [];

  // ── Login ─────────────────────────────────────────────────────
  let loginRes;
  try {
    loginRes = await axios.post(`${base}/api/method/login`, qs.stringify({ usr: username, pwd: password }), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });
  } catch (err) {
    const status = err.response ? err.response.status : null;
    if (status === 401 || status === 403) throw new Error('ERPNext login failed — check your username and password.');
    if (err.code === 'ECONNREFUSED' || err.code === 'ENOTFOUND') throw new Error('Cannot reach ERPNext server — check the URL.');
    throw new Error('ERPNext login error: ' + (err.message || 'Unknown'));
  }
  const cookie      = loginRes.headers['set-cookie'] ? loginRes.headers['set-cookie'].map(c => c.split(';')[0]).join('; ') : '';
  const authHeaders = { Cookie: cookie };
  const rt          = (reportType || 'full').toLowerCase();
  const emptyData   = { data: { data: [] } };

  // ── Helper: run a frappe query_report ─────────────────────────
  // Frappe expects filters as a JSON string inside the POST body
  async function runReport(reportName, filters, label) {
    try {
      const res = await axios.post(`${base}/api/method/frappe.desk.query_report.run`, {
        report_name: reportName,
        filters: JSON.stringify(filters),   // ← must be stringified
        are_default_filters: 0,
      }, { headers: { ...authHeaders, 'Content-Type': 'application/json' }, timeout: 60000 });
      const result = res.data.message;
      if (!result) { warnings.push(`${label}: empty response`); return []; }
      // result.result is array of row arrays; columns is array of {fieldname} objects
      const cols = (result.columns || []).map(c => typeof c === 'object' ? (c.fieldname || c.label) : c);
      const rows = (result.result || []).filter(row => Array.isArray(row)); // skip subtotals/group rows
      return rows.map(row => {
        const obj = {};
        cols.forEach((col, i) => { obj[col] = row[i]; });
        return obj;
      });
    } catch (err) {
      warnings.push(`${label}: ${err.message}`);
      return [];
    }
  }

  // ── Helper: fetch item brands from Item master ─────────────────
  async function fetchItemBrands() {
    const items = [];
    let start = 0;
    while (true) {
      try {
        const res = await axios.get(
          `${base}/api/resource/Item?fields=["name","item_name","brand","item_group"]&limit=500&limit_start=${start}&filters=[["Item","disabled","=",0]]`,
          { headers: authHeaders, timeout: 30000 }
        );
        const batch = res.data.data || [];
        items.push(...batch);
        if (batch.length < 500) break;
        start += 500;
      } catch (err) {
        warnings.push(`Item brands (page ${start}): ${err.message}`);
        break;
      }
    }
    const map = {};
    items.forEach(it => { map[it.name] = { brand: (it.brand || 'No Brand').trim(), category: (it.item_group || 'Uncategorised').trim(), item_name: (it.item_name || it.name || '').trim() }; });
    return map;
  }

  // ── SALES REPORT ─────────────────────────────────────────────
  // Mirrors generate_sales_report.js from the office PC exactly.
  // Uses Item-wise Sales Register report + salesperson lookup per invoice.
  if (rt === 'sales') {
    const now = new Date();
    const yr       = fiscalYear ? parseInt(fiscalYear) : now.getFullYear();
    const fromDate = `${yr}-01-01`;
    const toDate   = yr < now.getFullYear() ? `${yr}-12-31` : now.toISOString().slice(0, 10);

    // 1. Fetch submitted Sales Invoices in date range
    const invoices = [];
    let invStart = 0;
    while (true) {
      try {
        const res = await axios.get(
          `${base}/api/resource/Sales Invoice?fields=["name","posting_date","customer","customer_name","territory"]&filters=[["Sales Invoice","docstatus","=",1],["Sales Invoice","posting_date",">=","${fromDate}"],["Sales Invoice","posting_date","<=","${toDate}"]]&limit=500&limit_start=${invStart}`,
          { headers: authHeaders, timeout: 30000 }
        );
        const batch = res.data.data || [];
        invoices.push(...batch);
        if (batch.length < 500) break;
        invStart += 500;
      } catch (err) {
        warnings.push(`Sales Invoices fetch: ${err.message}`);
        break;
      }
    }

    // 2. Fetch each invoice individually — gets items + sales_team in one call,
    //    avoids needing "Sales Invoice Item" doctype read permissions
    const rawRows = [];
    const CONCURRENT = 15;
    for (let i = 0; i < invoices.length; i += CONCURRENT) {
      const batch = invoices.slice(i, i + CONCURRENT);
      const results = await Promise.all(batch.map(async (inv) => {
        try {
          const res = await axios.get(
            `${base}/api/resource/Sales%20Invoice/${encodeURIComponent(inv.name)}`,
            { headers: authHeaders, timeout: 20000 }
          );
          const doc = res.data.data || {};
          const items = doc.items || [];
          const st = (doc.sales_team || [])[0];
          const salesPerson = st ? (st.sales_person || '') : '';
          return items.map(item => ({
            parent       : inv.name,
            item_code    : item.item_code    || '',
            item_name    : item.item_name    || '',
            stock_qty    : item.stock_qty    || item.qty || 0,
            qty          : item.qty          || 0,
            rate         : item.rate         || 0,
            amount       : item.amount       || 0,
            posting_date : inv.posting_date  || '',
            customer     : inv.customer      || '',
            customer_name: inv.customer_name || '',
            territory    : doc.territory     || inv.territory || '',
            sales_person : salesPerson,
          }));
        } catch (err) {
          warnings.push(`Invoice ${inv.name}: ${err.message}`);
          return [];
        }
      }));
      results.forEach(items => rawRows.push(...items));
    }

    // 3. Fetch item brands (brand/category lookup)
    const [brandMap] = await Promise.all([
      fetchItemBrands(),
    ]);

    // 4. Merge into rows — same shape as office PC sales_items.json
    const toMonthKey = s => (s && s.length >= 7) ? s.substring(0, 7) : 'Unknown';
    const rows = rawRows.map(r => {
      const invName = r.parent || '';
      const info = brandMap[r.item_code] || { brand: 'No Brand', category: 'Uncategorised' };
      return {
        date       : toStdDate(r.posting_date || ''),
        month_key  : toMonthKey(r.posting_date || ''),
        invoice_no : invName,
        customer   : (r.customer_name || r.customer || 'Unknown').trim(),
        territory  : (r.territory || 'Unknown').trim(),
        salesperson: (r.sales_person || 'Unassigned').trim(),
        item_code  : r.item_code || '',
        item_name  : r.item_name || r.item_code || '',
        brand      : info.brand,
        category   : info.category,
        qty        : parseFloat(r.stock_qty || r.qty) || 0,
        rate       : parseFloat(r.rate) || 0,
        amount     : parseFloat(r.amount) || 0,
      };
    }).filter(r => r.item_code);

    rows.sort((a, b) => (a.date < b.date ? -1 : 1));

    // 4. Derive invoice summary list from rows
    const invAgg = {};
    rows.forEach(r => {
      if (!invAgg[r.invoice_no]) invAgg[r.invoice_no] = { name: r.invoice_no, customer: r.customer, posting_date: r.date, salesperson: r.salesperson, total: 0 };
      invAgg[r.invoice_no].total += r.amount;
    });
    const invoiceSummary = Object.values(invAgg);

    return { crmName: 'ERPNext', reportType: 'sales', warnings, rows, invoices: invoiceSummary, periodYear: yr, toDate, fromDate };
  }

  // ── STOCK REPORT ─────────────────────────────────────────────
  // Uses direct REST API (/api/resource/Bin) — reliable, no report filters needed.
  if (rt === 'stock') {

    // Fetch all Bin records (warehouse stock per item) via direct REST API
    const binFields = ['item_code','warehouse','actual_qty','stock_uom','valuation_rate','stock_value','reserved_qty','ordered_qty','indented_qty'];
    const bins = [];
    let binStart = 0;
    while (true) {
      try {
        const res = await axios.get(
          `${base}/api/resource/Bin?fields=${JSON.stringify(binFields)}&limit=500&limit_start=${binStart}&filters=[["Bin","actual_qty","!=",0]]`,
          { headers: authHeaders, timeout: 30000 }
        );
        const batch = res.data.data || [];
        bins.push(...batch);
        if (batch.length < 500) break;
        binStart += 500;
      } catch (err) {
        warnings.push(`Bin fetch (page ${binStart}): ${err.message}`);
        break;
      }
    }

    const brandMap = await fetchItemBrands();

    // Map Bin records to same shape as office PC stock_balance_raw rows
    const allRows = bins.map(r => ({
      Brand        : (brandMap[r.item_code] || {}).brand    || 'No Brand',
      Category     : (brandMap[r.item_code] || {}).category || 'Uncategorised',
      Item_Code    : r.item_code   || '',
      Item_Name    : (brandMap[r.item_code] || {}).item_name || r.item_code || '',
      Warehouse    : r.warehouse   || '',
      UOM          : r.stock_uom   || '',
      Opening_Qty  : 0,
      In_Qty       : 0,
      Out_Qty      : 0,
      Balance_Qty  : parseFloat(r.actual_qty)      || 0,
      Val_Rate     : parseFloat(r.valuation_rate)  || 0,
      Balance_Value: parseFloat(r.stock_value)     || 0,
      Reserved_Stock: parseFloat(r.reserved_qty)   || 0,
    })).filter(r => r.Item_Code)
      .sort((a, b) => a.Brand.localeCompare(b.Brand) || a.Category.localeCompare(b.Category) || a.Item_Code.localeCompare(b.Item_Code));

    // Aggregate summaries — same as office PC
    const brandAgg = {}, catAgg = {};
    for (const r of allRows) {
      if (!brandAgg[r.Brand]) brandAgg[r.Brand] = { brand: r.Brand, items: 0, qty: 0, value: 0 };
      brandAgg[r.Brand].items++; brandAgg[r.Brand].qty += r.Balance_Qty; brandAgg[r.Brand].value += r.Balance_Value;
      if (!catAgg[r.Category]) catAgg[r.Category] = { cat: r.Category, items: 0, qty: 0, value: 0 };
      catAgg[r.Category].items++; catAgg[r.Category].qty += r.Balance_Qty; catAgg[r.Category].value += r.Balance_Value;
    }
    const brandRows = Object.values(brandAgg).sort((a, b) => a.brand.localeCompare(b.brand));
    const catRows   = Object.values(catAgg).sort((a, b) => a.cat.localeCompare(b.cat));
    const stats = {
      totalItems : allRows.length,
      totalQty   : allRows.reduce((s, r) => s + r.Balance_Qty,   0),
      totalValue : allRows.reduce((s, r) => s + r.Balance_Value, 0),
      brands     : brandRows.length,
      categories : catRows.length,
      zeroStock  : allRows.filter(r => r.Balance_Qty <= 0).length,
      brandRows,
    };

    return { crmName: 'ERPNext', reportType: 'stock', warnings, allRows, brandRows, catRows, stats };
  }

  // ── AUDIT REPORT ─────────────────────────────────────────────
  // Mirrors audit_sales_data.js from the office PC exactly.
  // Runs 6 analytical checks on the sales data.
  if (rt === 'audit') {
    const now = new Date();
    const yr       = fiscalYear ? parseInt(fiscalYear) : now.getFullYear();
    const fromDate = `${yr}-01-01`;
    const toDate   = yr < now.getFullYear() ? `${yr}-12-31` : now.toISOString().slice(0, 10);

    // Fetch submitted Sales Invoices in date range, then items in chunks
    const auditInvoices = [];
    let auInvStart = 0;
    while (true) {
      try {
        const res = await axios.get(
          `${base}/api/resource/Sales Invoice?fields=["name","posting_date","customer","customer_name"]&filters=[["Sales Invoice","docstatus","=",1],["Sales Invoice","posting_date",">=","${fromDate}"],["Sales Invoice","posting_date","<=","${toDate}"]]&limit=500&limit_start=${auInvStart}`,
          { headers: authHeaders, timeout: 30000 }
        );
        const batch = res.data.data || [];
        auditInvoices.push(...batch);
        if (batch.length < 500) break;
        auInvStart += 500;
      } catch (err) { warnings.push(`Audit invoices: ${err.message}`); break; }
    }
    // Fetch each audit invoice individually — gets items + sales_team without needing
    // Sales Invoice Item doctype permissions
    const auditRaw = [];
    const AUDIT_CONCURRENT = 15;
    for (let i = 0; i < auditInvoices.length; i += AUDIT_CONCURRENT) {
      const batch = auditInvoices.slice(i, i + AUDIT_CONCURRENT);
      const results = await Promise.all(batch.map(async (inv) => {
        try {
          const res = await axios.get(
            `${base}/api/resource/Sales%20Invoice/${encodeURIComponent(inv.name)}`,
            { headers: authHeaders, timeout: 20000 }
          );
          const doc = res.data.data || {};
          const docItems = doc.items || [];
          const st = (doc.sales_team || [])[0];
          const salesPerson = st ? (st.sales_person || '') : '';
          return docItems.map(item => ({
            parent       : inv.name,
            item_code    : item.item_code    || '',
            item_name    : item.item_name    || '',
            stock_qty    : item.stock_qty    || item.qty || 0,
            qty          : item.qty          || 0,
            rate         : item.rate         || 0,
            amount       : item.amount       || 0,
            posting_date : inv.posting_date  || '',
            customer_name: inv.customer_name || '',
            sales_person : salesPerson,
          }));
        } catch (err) {
          warnings.push(`Audit invoice ${inv.name}: ${err.message}`);
          return [];
        }
      }));
      results.forEach(items => auditRaw.push(...items));
    }

    const [auditRawRows, brandMap] = await Promise.all([
      Promise.resolve(auditRaw),
      fetchItemBrands(),
    ]);
    const rawRows = auditRawRows;

    // Build items array in same format as audit_sales_data.js expects
    const items = rawRows.map(r => {
      const info = brandMap[r.item_code] || { brand: 'No Brand', category: 'Uncategorised' };
      return {
        parent        : r.parent || '',
        posting_date  : toStdDate(r.posting_date || ''),
        customer_name : (r.customer_name || '').trim(),
        salesperson   : (r.sales_person || '').trim(),
        item_code     : r.item_code || '',
        item_name     : r.item_name || '',
        qty           : parseFloat(r.stock_qty || r.qty) || 0,
        rate          : parseFloat(r.rate)   || 0,
        amount        : parseFloat(r.amount) || 0,
        brand         : info.brand,
        category      : info.category,
      };
    }).filter(r => r.item_code);

    // Derive invoices list
    const invAgg = {};
    items.forEach(i => {
      if (!invAgg[i.parent]) invAgg[i.parent] = { name: i.parent, customer: i.customer_name, posting_date: i.posting_date };
    });
    const invoices = Object.values(invAgg);

    const totalRev = items.filter(i => i.amount > 0).reduce((s, i) => s + i.amount, 0);

    // ── Run the same 6 checks as audit_sales_data.js ─────────────
    function dayName(s) { return ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][new Date(s).getDay()]; }
    const findings = { testItems:[], negativeAmounts:[], priceVariance:[], unassignedSP:[], weekendInvoices:[], concentration:[] };

    // CHECK 1: Test items
    items.filter(i => (i.item_code||'').toLowerCase().includes('test') || (i.item_name||'').toLowerCase().includes('test')).forEach(i => {
      findings.testItems.push({ risk:'HIGH', invoice:i.parent, date:i.posting_date, customer:i.customer_name, salesperson:i.salesperson, item_code:i.item_code, item_name:i.item_name, amount:i.amount, note:'Test item found in a live sales invoice. Should not exist in production.' });
    });

    // CHECK 2: Negative amounts
    items.filter(i => i.amount < 0).forEach(i => {
      const risk = i.amount < -1000 ? 'HIGH' : 'MEDIUM';
      findings.negativeAmounts.push({ risk, invoice:i.parent, date:i.posting_date, customer:i.customer_name, salesperson:i.salesperson, item_code:i.item_code, item_name:i.item_name, amount:i.amount, note: risk==='HIGH' ? 'Large negative amount — verify this is a valid credit note, not an error.' : 'Negative line item — confirm credit note is approved and authorised.' });
    });

    // CHECK 3: Price variance (same item, >30% spread)
    const byItem = {};
    items.forEach(i => { if (!byItem[i.item_code]) byItem[i.item_code] = []; byItem[i.item_code].push(i); });
    Object.entries(byItem).forEach(([code, sales]) => {
      const rates = sales.map(s => s.rate).filter(r => r > 0);
      if (rates.length < 2) return;
      const minRate = Math.min(...rates), maxRate = Math.max(...rates);
      const variance = (maxRate - minRate) / minRate;
      if (variance < 0.30) return;
      const risk = variance > 2 ? 'HIGH' : variance > 0.5 ? 'MEDIUM' : 'LOW';
      const minSale = sales.find(s => s.rate === minRate), maxSale = sales.find(s => s.rate === maxRate);
      findings.priceVariance.push({ risk, item_code:code, item_name:sales[0].item_name, times_sold:sales.length, min_rate:minRate, max_rate:maxRate, variance_pct:Math.round(variance*100), lowest_to:minSale?minSale.customer_name:'', lowest_inv:minSale?minSale.parent:'', highest_to:maxSale?maxSale.customer_name:'', highest_inv:maxSale?maxSale.parent:'', note: variance>2 ? 'CRITICAL: Price difference exceeds 200%. Likely a data entry error.' : `Price varied by ${Math.round(variance*100)}% across customers. Check if discount was authorised.` });
    });
    findings.priceVariance.sort((a, b) => b.variance_pct - a.variance_pct);

    // CHECK 4: No salesperson
    const unassignedByInv = {};
    items.filter(i => !i.salesperson || i.salesperson === 'Unassigned').forEach(i => {
      if (!unassignedByInv[i.parent]) unassignedByInv[i.parent] = { date:i.posting_date, customer:i.customer_name, rev:0, items:0 };
      unassignedByInv[i.parent].rev += i.amount; unassignedByInv[i.parent].items += 1;
    });
    Object.entries(unassignedByInv).sort((a,b) => b[1].rev - a[1].rev).forEach(([inv, d]) => {
      findings.unassignedSP.push({ risk:d.rev>5000?'MEDIUM':'LOW', invoice:inv, date:d.date, customer:d.customer, line_items:d.items, revenue:d.rev, note:'No salesperson assigned. Commission and accountability cannot be tracked.' });
    });

    // CHECK 5: Weekend invoices (Fri=5, Sat=6 — Saudi weekend)
    const revByDate = {}, invsByDate = {};
    items.forEach(i => {
      const dt = i.posting_date; if (!dt) return;
      revByDate[dt] = (revByDate[dt]||0) + i.amount;
      if (!invsByDate[dt]) invsByDate[dt] = new Set();
      invsByDate[dt].add(i.parent);
    });
    const seenDates = new Set();
    invoices.forEach(inv => {
      if (!inv.posting_date) return;
      const d = new Date(inv.posting_date);
      if (d.getDay() !== 5 && d.getDay() !== 6) return;
      if (seenDates.has(inv.posting_date)) return;
      seenDates.add(inv.posting_date);
      const invRev = revByDate[inv.posting_date] || 0;
      if (invRev <= 0) return;
      const invCount = invsByDate[inv.posting_date] ? invsByDate[inv.posting_date].size : 0;
      findings.weekendInvoices.push({ risk:invRev>10000?'MEDIUM':'LOW', invoice:Array.from(invsByDate[inv.posting_date]||[]).join(', '), date:inv.posting_date, day:dayName(inv.posting_date), customer:inv.customer, inv_count:invCount, revenue:invRev, note:`${invCount} invoice(s) dated on a weekend. Verify these were genuinely transacted on this date.` });
    });
    findings.weekendInvoices.sort((a, b) => b.revenue - a.revenue);

    // CHECK 6: Customer concentration (top 10)
    const byCust = {};
    items.filter(i => i.amount > 0).forEach(i => { byCust[i.customer_name] = (byCust[i.customer_name]||0) + i.amount; });
    Object.entries(byCust).sort((a,b) => b[1]-a[1]).slice(0,10).forEach(([cust, rev]) => {
      const pct = totalRev > 0 ? (rev/totalRev*100) : 0;
      findings.concentration.push({ risk:pct>20?'HIGH':pct>10?'MEDIUM':'LOW', customer:cust, revenue:rev, pct:parseFloat(pct.toFixed(1)), note:pct>20?'Over 20% of total revenue from one customer — high dependency risk.':pct>10?'Over 10% of total revenue from one customer — monitor closely.':'Top 10 customer by revenue — informational.' });
    });

    // Summary
    const unassignedTotalRev = Object.values(unassignedByInv).reduce((s,d)=>s+d.rev,0);
    const negTotal = items.filter(i=>i.amount<0).reduce((s,i)=>s+i.amount,0);
    findings.summary = [
      { category:'🔴  Test Items in Live Invoices',   count:findings.testItems.length,       risk:findings.testItems.length>0?'HIGH':'CLEAR', detail:'Test/dummy items found in submitted sales invoices. These should be deleted.' },
      { category:'🔴  Negative Amount Line Items',    count:findings.negativeAmounts.length,  risk:findings.negativeAmounts.length>0?'HIGH':'CLEAR', detail:`Total negative value: ${Math.abs(negTotal).toFixed(2)}. Verify each is an authorised credit note.` },
      { category:'🟡  Price Inconsistency',           count:findings.priceVariance.length,    risk:findings.priceVariance.some(f=>f.risk==='HIGH')?'HIGH':'MEDIUM', detail:`${findings.priceVariance.length} products sold at significantly different prices to different customers.` },
      { category:'🟡  Revenue Without Salesperson',   count:findings.unassignedSP.length,     risk:'MEDIUM', detail:`${unassignedTotalRev.toFixed(2)} revenue across ${findings.unassignedSP.length} invoices has no salesperson assigned.` },
      { category:'🟡  Weekend Invoices',              count:findings.weekendInvoices.length,  risk:findings.weekendInvoices.length>10?'MEDIUM':'LOW', detail:`${findings.weekendInvoices.length} invoices dated on Friday or Saturday.` },
      { category:'🔵  Customer Concentration',        count:findings.concentration.length,    risk:'INFO', detail:'Top 10 customers shown. Monitor if any single customer exceeds 20% of revenue.' },
    ];

    return { crmName:'ERPNext', reportType:'audit', warnings, items, invoices, findings, totalRev, periodYear: yr };
  }

  // ── FULL CRM REPORT (default) ─────────────────────────────────
  const fullWarnings = [];
  const emptyList = { data: { data: [] } };
  const [leadsRes, oppsRes, contactsRes, tasksRes] = await Promise.all([
    safeGet('Leads',    axios.get(`${base}/api/resource/Lead?fields=["name","lead_name","email_id","mobile_no","company_name","status","modified"]&limit=100`, { headers: authHeaders }), emptyList, fullWarnings),
    safeGet('Opportunities', axios.get(`${base}/api/resource/Opportunity?fields=["name","opportunity_from","opportunity_type","status","expected_closing","opportunity_amount","probability"]&limit=100`, { headers: authHeaders }), emptyList, fullWarnings),
    safeGet('Contacts', axios.get(`${base}/api/resource/Contact?fields=["name","first_name","last_name","email_id","mobile_no","company_name","designation","modified"]&limit=100`, { headers: authHeaders }), emptyList, fullWarnings),
    safeGet('Tasks',    axios.get(`${base}/api/resource/Task?fields=["name","subject","status","exp_end_date","task_weight"]&limit=100`, { headers: authHeaders }), emptyList, fullWarnings),
  ]);

  const leads    = leadsRes.data.data    || [];
  const opps     = oppsRes.data.data     || [];
  const contacts = contactsRes.data.data || [];
  const tasks    = tasksRes.data.data    || [];

  const pipeline = opps.map(o => ({
    name: o.opportunity_from||'', amount: parseFloat(o.opportunity_amount)||0,
    stage: o.status||'', closeDate: o.expected_closing||'',
    probability: parseFloat(o.probability)||0, pipeline: o.opportunity_type||'',
  }));
  const leadList = leads.map(l => ({
    name: l.lead_name||'', email: l.email_id||'', phone: l.mobile_no||'',
    company: l.company_name||'', status: l.status||'', lastActivity: l.modified||'',
  }));
  const contactList = contacts.map(c => ({
    name: `${c.first_name||''} ${c.last_name||''}`.trim(),
    email: c.email_id||'', phone: c.mobile_no||'',
    company: c.company_name||'', title: c.designation||'',
    lastActivity: c.modified||'', status: 'Active',
  }));
  const actList = tasks.map(t => ({
    subject: t.subject||'', status: t.status||'', dueDate: t.exp_end_date||'', type: 'Task',
  }));
  const wonDeals  = opps.filter(o=>o.status==='Won').map(o=>({name:o.opportunity_from,amount:parseFloat(o.opportunity_amount)||0,result:'Won',closeDate:o.expected_closing}));
  const lostDeals = opps.filter(o=>o.status==='Lost').map(o=>({name:o.opportunity_from,amount:parseFloat(o.opportunity_amount)||0,result:'Lost',closeDate:o.expected_closing}));
  const totalPipeline = pipeline.reduce((s,d)=>s+d.amount,0);
  const openDeals     = pipeline.filter(d=>!['Won','Lost'].includes(d.stage)).length;
  const closingSoon   = pipeline.filter(d=>{ if(!d.closeDate)return false; const days=(new Date(d.closeDate)-new Date())/86400000; return days>=0&&days<=30;}).length;
  const wonThisMonth  = wonDeals.filter(d=>{ if(!d.closeDate)return false; const cd=new Date(d.closeDate); const now=new Date(); return cd.getMonth()===now.getMonth()&&cd.getFullYear()===now.getFullYear();}).length;

  return { crmName:'ERPNext', reportType:'full', warnings:[...warnings,...fullWarnings], kpis:{totalPipeline,openDeals,closingSoon,wonThisMonth}, pipeline, leads:leadList, contacts:contactList, activities:actList, wonLost:[...wonDeals,...lostDeals] };
}

async function fetchSalesforce(instanceUrl, accessToken, reportType) {
  const base     = instanceUrl.replace(/\/$/, '');
  const headers  = { Authorization: `Bearer ${accessToken}` };
  const rt       = (reportType || 'full').toLowerCase();
  const warnings = [];
  const query    = (soql) => axios.get(`${base}/services/data/v58.0/query?q=${encodeURIComponent(soql)}`, { headers });

  const needDeals    = ['full','pipeline'].includes(rt);
  const needContacts = ['full','contacts'].includes(rt);
  const needLeads    = ['full','leads'].includes(rt);

  const emptyQ = { data: { records: [] } };
  const [dealsRes, contactsRes, leadsRes] = await Promise.all([
    needDeals    ? safeGet('Opportunities', query('SELECT Id,Name,Amount,StageName,CloseDate,Probability,Type FROM Opportunity LIMIT 100'), emptyQ, warnings) : Promise.resolve(emptyQ),
    needContacts ? safeGet('Contacts',      query('SELECT Id,FirstName,LastName,Email,Phone,Account.Name,Title,LastActivityDate FROM Contact LIMIT 100'), emptyQ, warnings) : Promise.resolve(emptyQ),
    needLeads    ? safeGet('Leads',         query('SELECT Id,FirstName,LastName,Email,Phone,Company,Status,LastActivityDate FROM Lead LIMIT 100'), emptyQ, warnings) : Promise.resolve(emptyQ),
  ]);

  const deals    = dealsRes.data.records    || [];
  const contacts = contactsRes.data.records || [];
  const leads    = leadsRes.data.records    || [];

  const pipeline    = deals.map(d => ({ name:d.Name||'', amount:parseFloat(d.Amount)||0, stage:d.StageName||'', closeDate:d.CloseDate||'', probability:parseFloat(d.Probability)||0, pipeline:d.Type||'' }));
  const contactList = contacts.map(c => ({ name:`${c.FirstName||''} ${c.LastName||''}`.trim(), email:c.Email||'', phone:c.Phone||'', company:c.Account?c.Account.Name:'', title:c.Title||'', lastActivity:c.LastActivityDate||'', status:'Active' }));
  const leadList    = leads.map(l => ({ name:`${l.FirstName||''} ${l.LastName||''}`.trim(), email:l.Email||'', phone:l.Phone||'', company:l.Company||'', status:l.Status||'', lastActivity:l.LastActivityDate||'' }));
  const wonDeals    = deals.filter(d=>d.StageName==='Closed Won').map(d=>({ name:d.Name, amount:parseFloat(d.Amount)||0, result:'Won', closeDate:d.CloseDate }));
  const lostDeals   = deals.filter(d=>d.StageName==='Closed Lost').map(d=>({ name:d.Name, amount:parseFloat(d.Amount)||0, result:'Lost', closeDate:d.CloseDate }));

  const totalPipeline = pipeline.reduce((s,d)=>s+d.amount,0);
  const openDeals     = pipeline.filter(d=>!['Closed Won','Closed Lost'].includes(d.stage)).length;
  const closingSoon   = pipeline.filter(d=>{ if(!d.closeDate)return false; const days=(new Date(d.closeDate)-new Date())/86400000; return days>=0&&days<=30; }).length;
  const wonThisMonth  = wonDeals.filter(d=>{ if(!d.closeDate)return false; const cd=new Date(d.closeDate); const now=new Date(); return cd.getMonth()===now.getMonth()&&cd.getFullYear()===now.getFullYear(); }).length;

  return { crmName:'Salesforce', reportType:rt, warnings, kpis:{totalPipeline,openDeals,closingSoon,wonThisMonth,totalLeads:leads.length,totalContacts:contacts.length}, pipeline, leads:leadList, contacts:contactList, activities:[], wonLost:[...wonDeals,...lostDeals] };
}

// ── ODOO FETCHER ──────────────────────────────────────────────
// Uses Odoo JSON-RPC (simpler than XML-RPC, no extra dependency needed)
// Endpoint: /web/dataset/call_kw  — works on Odoo 14, 15, 16, 17

async function odooCall(base, cookie, model, method, args, kwargs = {}) {
  const payload = {
    jsonrpc: '2.0', method: 'call', id: Math.floor(Math.random() * 9999),
    params: { model, method, args, kwargs: { ...kwargs, context: {} } },
  };
  const res = await axios.post(`${base}/web/dataset/call_kw`, payload, {
    headers: { 'Content-Type': 'application/json', 'Cookie': cookie },
    timeout: 30000,
  });
  if (res.data.error) throw new Error(`Odoo API error: ${res.data.error.data?.message || res.data.error.message}`);
  return res.data.result;
}

async function fetchOdoo(url, db, username, password, reportType, fiscalYear) {
  const base = url.replace(/\/$/, '');
  const warnings = [];
  const rt = (reportType || 'full').toLowerCase();

  // ── Step 1: Build database name candidates ────────────────────
  // Try user-supplied db first, then common Odoo naming patterns
  const hostname = (() => { try { return new URL(base).hostname; } catch { return ''; } })();
  const subdomain = hostname.split('.')[0]; // e.g. "mycompany" from "mycompany.odoo.com"
  const dbCandidates = [];
  if (db && db.trim()) dbCandidates.push(db.trim());       // user-supplied: highest priority
  dbCandidates.push(subdomain);                             // e.g. "mycompany"
  dbCandidates.push(hostname.replace(/\./g, '-'));          // e.g. "mycompany-odoo-com"
  dbCandidates.push(hostname.replace(/\.[^.]+$/, '').replace(/\./g, '_')); // e.g. "mycompany_odoo"

  // ── Step 2: Also try to fetch available databases ─────────────
  try {
    const dbListRes = await axios.post(`${base}/web/database/list`, {
      jsonrpc: '2.0', method: 'call', id: 1, params: {},
    }, { headers: { 'Content-Type': 'application/json' }, timeout: 8000 });
    const dbs = dbListRes.data.result || [];
    if (dbs.length > 0) dbCandidates.unshift(dbs[0]); // prepend first available DB
  } catch (_) { /* database list not exposed — that's fine, continue with candidates */ }

  // ── Step 3: Try each database candidate until login succeeds ──
  let uid, sessionCookie, effectiveDb;
  const loginErrors = [];

  for (const candidate of dbCandidates) {
    if (!candidate) continue;
    try {
      const loginRes = await axios.post(`${base}/web/session/authenticate`, {
        jsonrpc: '2.0', method: 'call', id: 1,
        params: { db: candidate, login: username, password },
      }, { headers: { 'Content-Type': 'application/json' }, timeout: 20000 });

      if (loginRes.data.error) {
        const msg = loginRes.data.error.data?.message || loginRes.data.error.message || '';
        if (msg.includes('Database not found') || msg.includes('database') || msg.includes('does not exist')) {
          loginErrors.push(`db "${candidate}": not found`);
          continue; // try next candidate
        }
        if (msg.includes('Access Denied') || msg.includes('credentials')) {
          throw new Error('Wrong username or password for database "' + candidate + '".');
        }
        loginErrors.push(`db "${candidate}": ${msg}`);
        continue;
      }

      const tempUid = loginRes.data.result && loginRes.data.result.uid;
      if (!tempUid) {
        loginErrors.push(`db "${candidate}": wrong credentials`);
        continue;
      }

      // Success
      uid          = tempUid;
      effectiveDb  = candidate;
      const rawCookies = loginRes.headers['set-cookie'] || [];
      sessionCookie = rawCookies.map(c => c.split(';')[0]).join('; ');
      break;

    } catch (err) {
      if (err.code === 'ECONNREFUSED' || err.code === 'ENOTFOUND') {
        throw new Error('Cannot reach Odoo server — check the URL.');
      }
      if (err.message.startsWith('Wrong username')) throw err;
      loginErrors.push(`db "${candidate}": ${err.message}`);
    }
  }

  if (!uid) {
    throw new Error(
      'Odoo login failed. Could not find the database automatically. ' +
      'Please enter your database name in the "Database Name" field. ' +
      'Tried: ' + dbCandidates.filter(Boolean).join(', ') + '. ' +
      'Tip: In Odoo, go to Settings → Technical → Database to find the name.'
    );
  }

  // ── Helper: safe Odoo call ────────────────────────────────────
  async function safeOdoo(label, model, method, args, kwargs) {
    try {
      return await odooCall(base, sessionCookie, model, method, args, kwargs);
    } catch (err) {
      warnings.push(`${label}: ${err.message}`);
      return [];
    }
  }

  // ── Step 3: Fetch data based on reportType ────────────────────
  const needPipeline  = ['full','pipeline'].includes(rt);
  const needContacts  = ['full','leads'].includes(rt);
  const needLeads     = ['full','leads'].includes(rt);
  const needActivities= ['full','activities'].includes(rt);

  // Year filter — applied to pipeline (close date) and activities (due date)
  const odooYr      = fiscalYear ? parseInt(fiscalYear) : null;
  const yrFrom      = odooYr ? `${odooYr}-01-01` : null;
  const yrTo        = odooYr ? `${odooYr}-12-31` : null;
  const pipelineDomain = odooYr
    ? [['active','=',true],['date_deadline','>=',yrFrom],['date_deadline','<=',yrTo]]
    : [['active','=',true]];
  const activitiesDomain = odooYr
    ? [['res_model','in',['crm.lead','sale.order']],['date_deadline','>=',yrFrom],['date_deadline','<=',yrTo]]
    : [['res_model','in',['crm.lead','sale.order']]];

  // CRM Leads / Opportunities
  // NOTE: Odoo 16+ removed the 'type' field from crm.lead — all records are just CRM entries.
  // We fetch ALL active crm.lead records and split by 'type' field if present.
  const allLeads = needPipeline || needLeads
    ? await safeOdoo('CRM Records', 'crm.lead', 'search_read',
        [needPipeline && odooYr ? pipelineDomain : [['active','=',true]]],
        { fields: ['name','partner_name','email_from','phone','planned_revenue','stage_id','probability','date_deadline','user_id','kanban_state','type','create_date','source_id'], limit: 500 }
      )
    : [];

  // Split: if 'type' field exists use it; otherwise treat all as opportunities (Odoo 16+)
  const hasTypeField = allLeads.length > 0 && allLeads[0].hasOwnProperty('type') && allLeads[0].type;
  const opps = needPipeline
    ? (hasTypeField ? allLeads.filter(r => r.type === 'opportunity') : allLeads)
    : [];
  const leads = needLeads
    ? (hasTypeField ? allLeads.filter(r => r.type === 'lead') : [])
    : [];

  const [contacts, activities] = await Promise.all([
    needContacts ? safeOdoo('Contacts', 'res.partner', 'search_read',
      [[['is_company','=',false],['active','=',true]]],
      { fields: ['name','email','phone','mobile','company_name','job_position','create_date'], limit: 200 }
    ) : Promise.resolve([]),

    needActivities ? safeOdoo('Activities', 'mail.activity', 'search_read',
      [activitiesDomain],
      { fields: ['activity_type_id','summary','note','date_deadline','state','user_id','res_name'], limit: 200 }
    ) : Promise.resolve([]),
  ]);

  // ── Map to standard shapes ────────────────────────────────────
  const pipeline = opps.map(o => ({
    name:        o.name || '',
    amount:      parseFloat(o.planned_revenue) || 0,
    stage:       Array.isArray(o.stage_id) ? o.stage_id[1] : (o.stage_id || ''),
    closeDate:   toStdDate(o.date_deadline || ''),
    probability: parseFloat(o.probability) || 0,
    pipeline:    Array.isArray(o.user_id) ? o.user_id[1] : '',
  }));

  const leadList = leads.map(l => ({
    name:        l.name || l.partner_name || '',
    email:       l.email_from || '',
    phone:       l.phone || '',
    company:     l.partner_name || '',
    status:      Array.isArray(l.stage_id) ? l.stage_id[1] : '',
    source:      Array.isArray(l.source_id) ? l.source_id[1] : (l.source_id || 'Direct'),
    lastActivity: toStdDate(l.create_date || ''),
  }));

  const contactList = contacts.map(c => ({
    name:        c.name || '',
    email:       c.email || '',
    phone:       c.phone || c.mobile || '',
    company:     c.company_name || '',
    title:       c.job_position || '',
    lastActivity: toStdDate(c.create_date || ''),
    status:      'Active',
  }));

  const activityList = activities.map(a => ({
    subject:  a.summary || (Array.isArray(a.activity_type_id) ? a.activity_type_id[1] : '') || '',
    status:   a.state || '',
    dueDate:  toStdDate(a.date_deadline || ''),
    type:     Array.isArray(a.activity_type_id) ? a.activity_type_id[1] : '',
  }));

  const wonOpps  = opps.filter(o => o.kanban_state === 'done' || (Array.isArray(o.stage_id) && o.stage_id[1] && o.stage_id[1].toLowerCase().includes('won')));
  const lostOpps = opps.filter(o => Array.isArray(o.stage_id) && o.stage_id[1] && o.stage_id[1].toLowerCase().includes('lost'));
  const wonLost  = [
    ...wonOpps.map(o  => ({ name: o.name, amount: parseFloat(o.planned_revenue)||0, result: 'Won',  closeDate: toStdDate(o.date_deadline) })),
    ...lostOpps.map(o => ({ name: o.name, amount: parseFloat(o.planned_revenue)||0, result: 'Lost', closeDate: toStdDate(o.date_deadline) })),
  ];

  const totalPipeline = pipeline.reduce((s, d) => s + d.amount, 0);
  const openDeals     = pipeline.filter(d => !wonOpps.find(o => o.name === d.name) && !lostOpps.find(o => o.name === d.name)).length;
  const closingSoon   = pipeline.filter(d => {
    if (!d.closeDate) return false;
    const days = (new Date(d.closeDate) - new Date()) / 86400000;
    return days >= 0 && days <= 30;
  }).length;
  const wonThisMonth  = wonOpps.filter(o => {
    if (!o.date_deadline) return false;
    const cd = new Date(o.date_deadline); const now = new Date();
    return cd.getMonth() === now.getMonth() && cd.getFullYear() === now.getFullYear();
  }).length;

  return {
    crmName: 'Odoo',
    reportType: rt,
    periodYear: odooYr,
    warnings,
    kpis: { totalPipeline, openDeals, closingSoon, wonThisMonth, totalLeads: leadList.length, totalContacts: contactList.length, totalActivities: activityList.length },
    pipeline,
    leads: leadList,
    contacts: contactList,
    activities: activityList,
    wonLost,
    rawData: allLeads,         // unprocessed Odoo API response — used for Raw Data sheet
    rawContacts: contacts,     // unprocessed contacts from Odoo
  };
}

// ── EXCEL BUILDER ─────────────────────────────────────────────
async function buildExcel(data) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CRM Data Extractor — Developed by Arun Chiramal';
  wb.created = new Date();

  if (data.reportType === 'sales') {
    // 11 sheets — mirrors generate_sales_report.js from office PC exactly
    erpBuildSalesSummary(wb, data);
    erpBuildMonthWise(wb, data.rows);
    erpBuildProductWise(wb, data.rows);
    erpBuildBrandWise(wb, data.rows);
    erpBuildSalesperson(wb, data.rows);
    erpBuildCustomerWise(wb, data.rows);
    erpBuildPivot(wb, data.rows);
    erpBuildSPMonthWise(wb, data.rows);
    erpBuildSPProducts(wb, data.rows);
    erpBuildRawData(wb, data.rows);
    erpBuildSPDrilldown(wb, data.rows);
  } else if (data.reportType === 'stock') {
    // Summary + All Stock + By Brand + By Category + one tab per brand
    erpBuildStockSummary(wb, data);
    erpBuildAllStock(wb, data.allRows);
    erpBuildBrandSummary(wb, data.brandRows);
    erpBuildCatSummary(wb, data.catRows);
    const brandGroups = {};
    (data.allRows || []).forEach(r => {
      if (!brandGroups[r.Brand]) brandGroups[r.Brand] = [];
      brandGroups[r.Brand].push(r);
    });
    Object.keys(brandGroups).sort().forEach((brand, idx) => {
      erpBuildBrandSheet(wb, brand, brandGroups[brand], BRAND_COLORS[idx % BRAND_COLORS.length]);
    });
  } else if (data.reportType === 'audit') {
    // Summary + conditional check sheets
    erpBuildAuditSummary(wb, data);
    if ((data.findings.testItems||[]).length)       erpBuildTestItems(wb, data.findings);
    if ((data.findings.negativeAmounts||[]).length) erpBuildNegativeAmounts(wb, data.findings);
    if ((data.findings.priceVariance||[]).length)   erpBuildPriceVariance(wb, data.findings);
    if ((data.findings.unassignedSP||[]).length)    erpBuildUnassignedSP(wb, data.findings);
    if ((data.findings.weekendInvoices||[]).length) erpBuildWeekendInvoices(wb, data.findings);
    erpBuildConcentration(wb, data.findings);
  } else {
    const rt = (data.reportType || 'pipeline').toLowerCase();
    if (rt === 'pipeline') {
      // Sales Pipeline Report — 9 sheets
      buildDashboard(wb, data);
      buildPipeline(wb, data.pipeline);
      buildStageAnalysis(wb, data.pipeline);
      buildSalespersonSheet(wb, data.pipeline, data.wonLost);
      buildCustomerWiseSheet(wb, data.pipeline);
      buildMonthlyTrend(wb, data.pipeline, data.wonLost);
      buildForecastSheet(wb, data.pipeline);
      buildDealVelocity(wb, data.pipeline, data.wonLost);
      buildWonLost(wb, data.wonLost);
    } else if (rt === 'leads') {
      // Leads & Contacts Report — 4 sheets
      buildLeadsContactsDashboard(wb, data);
      buildLeads(wb, data.leads);
      buildLeadSourceAnalysis(wb, data.leads);
      buildContacts(wb, data.contacts);
    } else if (rt === 'activities') {
      // Activity Report — 4 sheets
      buildActivitiesDashboard(wb, data);
      buildActivities(wb, data.activities);
      buildOverdueActivities(wb, data.activities);
      buildActivitiesByUser(wb, data.activities);
    } else {
      // Fallback full — kept for backward compat
      buildDashboard(wb, data);
      buildPipeline(wb, data.pipeline);
      buildStageAnalysis(wb, data.pipeline);
      buildSalespersonSheet(wb, data.pipeline, data.wonLost);
      buildCustomerWiseSheet(wb, data.pipeline);
      buildMonthlyTrend(wb, data.pipeline, data.wonLost);
      buildForecastSheet(wb, data.pipeline);
      buildDealVelocity(wb, data.pipeline, data.wonLost);
      buildLeads(wb, data.leads);
      buildContacts(wb, data.contacts);
      buildActivities(wb, data.activities);
      buildWonLost(wb, data.wonLost);
      buildOdooRawData(wb, data.rawData, data.rawContacts);
    }
  }

  if (data.warnings && data.warnings.length > 0) {
    buildWarningsSheet(wb, data.warnings, data.crmName);
  }

  const buffer = await wb.xlsx.writeBuffer();
  return buffer;
}

// ── ERP SALES — 11 SHEETS (mirrors generate_sales_report.js) ─

function erpBuildSalesSummary(wb, data) {
  const rows = data.rows || [];
  const totalRev = rows.reduce((s,r)=>s+r.amount,0);
  const totalQty = rows.reduce((s,r)=>s+r.qty,0);
  const uniqueCustomers = new Set(rows.map(r=>r.customer)).size;
  const brands = [...new Set(rows.map(r=>r.brand))].filter(b=>b&&b!=='No Brand').length;
  const uniqueInvoices = new Set(rows.map(r=>r.invoice_no)).size;
  const months = [...new Set(rows.map(r=>r.month_key))].sort();
  const yr = data.periodYear;
  const now = new Date();
  const periodEnd = data.toDate ? toStdDate(data.toDate) : (yr && yr >= now.getFullYear() ? toStdDate(now) : toStdDate(`${yr}-12-31`));
  const periodLabel = yr ? `01 01 ${yr} – ${periodEnd}` : (months.length ? `${toMonthDisplay(months[0])} to ${toMonthDisplay(months[months.length-1])}` : fmtDate());

  const ws = wb.addWorksheet('📊 Summary', { properties:{ tabColor:{ argb:'FF'+C.darkBlue } } });
  ws.views=[{showGridLines:false}]; setColWidths(ws,[3,36,26,26,3]);
  ws.mergeCells('B2:D2');
  const t=ws.getCell('B2'); t.value='ERPNEXT — SALES REPORT';
  t.style={font:{bold:true,size:22,name:'Calibri',color:{argb:'FF'+C.headerFont}},fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF'+C.darkBlue}},alignment:{horizontal:'center',vertical:'middle'}};
  ws.getRow(2).height=48;
  ws.mergeCells('B3:D3');
  const s=ws.getCell('B3'); s.value=`Period: ${periodLabel}  |  Generated: ${fmtDate()}  |  Developed by Arun Chiramal`;
  s.style={font:{bold:false,size:10,name:'Calibri',color:{argb:'FF'+C.headerFont}},fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF'+C.midBlue}},alignment:{horizontal:'center',vertical:'middle'}};
  ws.getRow(3).height=22; ws.getRow(4).height=12;
  const kpis=[
    ['💰 Total Revenue',    totalRev.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}), C.darkBlue],
    ['🧾 Total Invoices',   uniqueInvoices.toString(),   C.midBlue],
    ['🏢 Unique Customers', uniqueCustomers.toString(),  C.teal],
    ['🏷️ Brands Sold',      brands.toString(),           C.purple],
    ['📦 Total Qty Sold',   totalQty.toLocaleString(),   C.orange],
    ['📅 Months of Data',   months.length.toString(),    '375623'],
  ];
  kpis.forEach(([label,value,bg],i)=>{
    const row=5+i; ws.mergeCells(`B${row}:C${row}`);
    const lC=ws.getCell(`B${row}`); lC.value=label;
    lC.style={font:{bold:true,size:11,name:'Calibri',color:{argb:'FF333333'}},fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFE9F0FB'}},alignment:{horizontal:'left',vertical:'middle',indent:1},border:thinBorder()};
    const vC=ws.getCell(`D${row}`); vC.value=value;
    vC.style={font:{bold:true,size:14,name:'Calibri',color:{argb:'FF'+C.headerFont}},fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF'+bg}},alignment:{horizontal:'center',vertical:'middle'},border:thinBorder()};
    ws.getRow(row).height=30;
  });
  ws.getRow(12).height=12;
  ws.mergeCells('B13:D13');
  const idx=ws.getCell('B13'); idx.value='📋  CONTAINS: Month-wise | Product-wise | Brand-wise | Salesperson | Customer-wise | Pivot | SP Month | SP Products | Raw Data | SP Drill-down';
  idx.style={font:{italic:true,size:10,name:'Calibri',color:{argb:'FF'+C.midBlue}},alignment:{horizontal:'center',vertical:'middle',wrapText:true}};
  ws.getRow(13).height=28;
}

function erpBuildMonthWise(wb, rows) {
  const agg={};
  rows.forEach(r=>{
    if(!agg[r.month_key]) agg[r.month_key]={month:r.month_key,invoices:new Set(),customers:new Set(),qty:0,amount:0};
    agg[r.month_key].invoices.add(r.invoice_no); agg[r.month_key].customers.add(r.customer);
    agg[r.month_key].qty+=r.qty; agg[r.month_key].amount+=r.amount;
  });
  const sorted=Object.values(agg).sort((a,b)=>a.month.localeCompare(b.month));
  const ws=wb.addWorksheet('📅 Month-wise',{properties:{tabColor:{argb:'FF'+C.midBlue}}});
  addSheetHeader(ws,'📅 Month-wise Revenue','Monthly Sales Breakdown',5);
  setColWidths(ws,[16,14,14,14,18]);
  const hRow=ws.getRow(4);
  ['Month','Invoices','Customers','Qty Sold','Revenue'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.midBlue);});
  hRow.height=20; freezeRow(ws,4);
  let totalRev=0;
  sorted.forEach((m,i)=>{
    const row=ws.getRow(5+i); const bg=i%2===0?C.white:C.altRow;
    row.getCell(1).value=toMonthDisplay(m.month); row.getCell(1).style=cellStyle(bg,true,'center');
    row.getCell(2).value=m.invoices.size; row.getCell(2).style=cellStyle(bg,false,'center');
    row.getCell(3).value=m.customers.size;row.getCell(3).style=cellStyle(bg,false,'center');
    row.getCell(4).value=m.qty;           row.getCell(4).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(5).value=m.amount;        row.getCell(5).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.height=18; totalRev+=m.amount;
  });
  const tot=ws.getRow(5+sorted.length); tot.height=22;
  [1,2,3,4,5].forEach(i=>{tot.getCell(i).style=headerStyle(C.darkBlue);});
  tot.getCell(1).value='TOTAL';
  tot.getCell(5).value=totalRev; tot.getCell(5).style={...headerStyle(C.darkBlue),numFmt:'#,##0.00'};
  if(!sorted.length){ws.getRow(5).getCell(1).value='No data.';}
}

function erpBuildProductWise(wb, rows) {
  const agg={};
  rows.forEach(r=>{
    const k=r.item_code;
    if(!agg[k]) agg[k]={item_code:k,item_name:r.item_name,brand:r.brand,category:r.category,qty:0,amount:0,invoices:new Set()};
    agg[k].qty+=r.qty; agg[k].amount+=r.amount; agg[k].invoices.add(r.invoice_no);
  });
  const sorted=Object.values(agg).sort((a,b)=>b.amount-a.amount);
  const ws=wb.addWorksheet('📦 Product-wise',{properties:{tabColor:{argb:'FF'+C.orange}}});
  addSheetHeader(ws,'📦 Product-wise Sales','Top Items by Revenue',7);
  setColWidths(ws,[18,30,16,18,12,14,18]);
  const hRow=ws.getRow(4);
  ['Item Code','Item Name','Brand','Category','Invoices','Qty Sold','Revenue'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.orange);});
  hRow.height=20; freezeRow(ws,4);
  sorted.forEach((p,i)=>{
    const row=ws.getRow(5+i); const bg=i%2===0?C.white:C.altRow;
    row.getCell(1).value=p.item_code;     row.getCell(1).style=cellStyle(bg,true);
    row.getCell(2).value=p.item_name;     row.getCell(2).style=cellStyle(bg);
    row.getCell(3).value=p.brand;         row.getCell(3).style=cellStyle(bg);
    row.getCell(4).value=p.category;      row.getCell(4).style=cellStyle(bg);
    row.getCell(5).value=p.invoices.size; row.getCell(5).style=cellStyle(bg,false,'center');
    row.getCell(6).value=p.qty;           row.getCell(6).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(7).value=p.amount;        row.getCell(7).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.height=18;
  });
  if(!sorted.length){ws.getRow(5).getCell(1).value='No data.';}
}

function erpBuildBrandWise(wb, rows) {
  const agg={};
  rows.forEach(r=>{
    const k=r.brand||'No Brand';
    if(!agg[k]) agg[k]={brand:k,products:new Set(),customers:new Set(),qty:0,amount:0};
    agg[k].products.add(r.item_code); agg[k].customers.add(r.customer);
    agg[k].qty+=r.qty; agg[k].amount+=r.amount;
  });
  const sorted=Object.values(agg).sort((a,b)=>b.amount-a.amount);
  const totalRev=sorted.reduce((s,b)=>s+b.amount,0);
  const ws=wb.addWorksheet('🏷️ Brand-wise',{properties:{tabColor:{argb:'FF'+C.purple}}});
  addSheetHeader(ws,'🏷️ Brand-wise Sales','Revenue & Quantity by Brand',7);
  setColWidths(ws,[22,12,12,14,18,12,12]);
  const hRow=ws.getRow(4);
  ['Brand','Products','Customers','Qty Sold','Revenue','% Share','Rank'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.purple);});
  hRow.height=20; freezeRow(ws,4);
  sorted.forEach((b,i)=>{
    const row=ws.getRow(5+i); const bg=i%2===0?C.white:C.altRow;
    const pct=totalRev>0?(b.amount/totalRev*100):0;
    row.getCell(1).value=b.brand;          row.getCell(1).style=cellStyle(bg,true);
    row.getCell(2).value=b.products.size;  row.getCell(2).style=cellStyle(bg,false,'center');
    row.getCell(3).value=b.customers.size; row.getCell(3).style=cellStyle(bg,false,'center');
    row.getCell(4).value=b.qty;            row.getCell(4).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(5).value=b.amount;         row.getCell(5).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(6).value=parseFloat(pct.toFixed(1)); row.getCell(6).style={...cellStyle(bg,false,'right'),numFmt:'0.0"%"'};
    row.getCell(7).value=i+1;              row.getCell(7).style=cellStyle(bg,false,'center');
    row.height=18;
  });
  if(!sorted.length){ws.getRow(5).getCell(1).value='No data.';}
}

function erpBuildSalesperson(wb, rows) {
  const agg={};
  rows.forEach(r=>{
    const k=r.salesperson||'Unassigned';
    if(!agg[k]) agg[k]={sp:k,invoices:new Set(),customers:new Set(),qty:0,amount:0};
    agg[k].invoices.add(r.invoice_no); agg[k].customers.add(r.customer);
    agg[k].qty+=r.qty; agg[k].amount+=r.amount;
  });
  const sorted=Object.values(agg).sort((a,b)=>b.amount-a.amount);
  const totalRev=sorted.reduce((s,sp)=>s+sp.amount,0);
  const ws=wb.addWorksheet('👤 Salesperson',{properties:{tabColor:{argb:'FF'+C.teal}}});
  addSheetHeader(ws,'👤 Salesperson Performance','Revenue by Sales Representative',7);
  setColWidths(ws,[28,12,14,14,18,12,12]);
  const hRow=ws.getRow(4);
  ['Salesperson','Invoices','Customers','Qty Sold','Revenue','% Share','Rank'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.teal);});
  hRow.height=20; freezeRow(ws,4);
  sorted.forEach((sp,i)=>{
    const row=ws.getRow(5+i);
    const bg=i===0?'FFD700':i===1?'C0C0C0':i===2?'CD7F32':i%2===0?C.white:C.altRow;
    const pct=totalRev>0?(sp.amount/totalRev*100):0;
    row.getCell(1).value=sp.sp;            row.getCell(1).style=cellStyle(bg,true);
    row.getCell(2).value=sp.invoices.size; row.getCell(2).style=cellStyle(bg,false,'center');
    row.getCell(3).value=sp.customers.size;row.getCell(3).style=cellStyle(bg,false,'center');
    row.getCell(4).value=sp.qty;           row.getCell(4).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(5).value=sp.amount;        row.getCell(5).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(6).value=parseFloat(pct.toFixed(1)); row.getCell(6).style={...cellStyle(bg,false,'right'),numFmt:'0.0"%"'};
    row.getCell(7).value=i+1;              row.getCell(7).style=cellStyle(bg,false,'center');
    row.height=18;
  });
  if(!sorted.length){ws.getRow(5).getCell(1).value='No data.';}
}

function erpBuildCustomerWise(wb, rows) {
  const agg={};
  rows.forEach(r=>{
    const k=r.customer;
    if(!agg[k]) agg[k]={customer:k,invoices:new Set(),products:new Set(),qty:0,amount:0};
    agg[k].invoices.add(r.invoice_no); agg[k].products.add(r.item_code);
    agg[k].qty+=r.qty; agg[k].amount+=r.amount;
  });
  const sorted=Object.values(agg).sort((a,b)=>b.amount-a.amount);
  const totalRev=sorted.reduce((s,c)=>s+c.amount,0);
  const ws=wb.addWorksheet('🏢 Customer-wise',{properties:{tabColor:{argb:'FF'+C.darkBlue}}});
  addSheetHeader(ws,'🏢 Customer-wise Revenue','Top Customers by Revenue',7);
  setColWidths(ws,[32,12,12,14,18,12,12]);
  const hRow=ws.getRow(4);
  ['Customer','Invoices','Products','Qty Bought','Revenue','% Share','Rank'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.darkBlue);});
  hRow.height=20; freezeRow(ws,4);
  sorted.forEach((c,i)=>{
    const row=ws.getRow(5+i); const bg=i%2===0?C.white:C.altRow;
    const pct=totalRev>0?(c.amount/totalRev*100):0;
    row.getCell(1).value=c.customer;       row.getCell(1).style=cellStyle(bg,true);
    row.getCell(2).value=c.invoices.size;  row.getCell(2).style=cellStyle(bg,false,'center');
    row.getCell(3).value=c.products.size;  row.getCell(3).style=cellStyle(bg,false,'center');
    row.getCell(4).value=c.qty;            row.getCell(4).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(5).value=c.amount;         row.getCell(5).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(6).value=parseFloat(pct.toFixed(1)); row.getCell(6).style={...cellStyle(pct>20?'FFE8D6':bg,false,'right'),numFmt:'0.0"%"'};
    row.getCell(7).value=i+1;              row.getCell(7).style=cellStyle(bg,false,'center');
    row.height=18;
  });
  if(!sorted.length){ws.getRow(5).getCell(1).value='No data.';}
}

function erpBuildPivot(wb, rows) {
  const months=[...new Set(rows.map(r=>r.month_key))].sort();
  const brands=[...new Set(rows.map(r=>r.brand||'No Brand'))].sort();
  const matrix={};
  rows.forEach(r=>{
    if(!matrix[r.month_key]) matrix[r.month_key]={};
    matrix[r.month_key][r.brand||'No Brand']=(matrix[r.month_key][r.brand||'No Brand']||0)+r.amount;
  });
  const ws=wb.addWorksheet('📈 Month×Brand Pivot',{properties:{tabColor:{argb:'FF'+C.midBlue}}});
  addSheetHeader(ws,'📈 Month × Brand Revenue Pivot','Revenue Matrix: Month vs Brand',brands.length+1);
  setColWidths(ws,[16,...brands.map(()=>16)]);
  const hRow=ws.getRow(4);
  hRow.getCell(1).value='Month'; hRow.getCell(1).style=headerStyle(C.darkBlue);
  brands.forEach((b,i)=>{hRow.getCell(i+2).value=b;hRow.getCell(i+2).style=headerStyle(C.midBlue);});
  hRow.height=22; freezeRow(ws,4);
  months.forEach((m,i)=>{
    const row=ws.getRow(5+i); const bg=i%2===0?C.white:C.altRow;
    row.getCell(1).value=toMonthDisplay(m); row.getCell(1).style=cellStyle(bg,true,'center');
    brands.forEach((b,j)=>{
      const v=(matrix[m]||{})[b]||0;
      row.getCell(j+2).value=v||null;
      row.getCell(j+2).style={...cellStyle(v>0?bg:'F3F4F6',false,'right'),numFmt:'#,##0'};
    });
    row.height=18;
  });
  if(!months.length){ws.getRow(5).getCell(1).value='No data.';}
}

function erpBuildSPMonthWise(wb, rows) {
  const agg={};
  rows.forEach(r=>{
    const k=`${r.salesperson||'Unassigned'}||${r.month_key}`;
    if(!agg[k]) agg[k]={sp:r.salesperson||'Unassigned',month:r.month_key,invoices:new Set(),amount:0};
    agg[k].invoices.add(r.invoice_no); agg[k].amount+=r.amount;
  });
  const sorted=Object.values(agg).sort((a,b)=>a.sp.localeCompare(b.sp)||a.month.localeCompare(b.month));
  const ws=wb.addWorksheet('👤 SP Month-wise',{properties:{tabColor:{argb:'FF'+C.teal}}});
  addSheetHeader(ws,'👤 SP × Month Revenue','Salesperson Performance by Month',4);
  setColWidths(ws,[28,16,12,18]);
  const hRow=ws.getRow(4);
  ['Salesperson','Month','Invoices','Revenue'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.teal);});
  hRow.height=20; freezeRow(ws,4);
  sorted.forEach((r,i)=>{
    const row=ws.getRow(5+i); const bg=i%2===0?C.white:C.altRow;
    row.getCell(1).value=r.sp;             row.getCell(1).style=cellStyle(bg,true);
    row.getCell(2).value=toMonthDisplay(r.month); row.getCell(2).style=cellStyle(bg,false,'center');
    row.getCell(3).value=r.invoices.size;  row.getCell(3).style=cellStyle(bg,false,'center');
    row.getCell(4).value=r.amount;         row.getCell(4).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.height=18;
  });
  if(!sorted.length){ws.getRow(5).getCell(1).value='No data.';}
}

function erpBuildSPProducts(wb, rows) {
  const agg={};
  rows.forEach(r=>{
    const k=`${r.salesperson||'Unassigned'}||${r.item_code}`;
    if(!agg[k]) agg[k]={sp:r.salesperson||'Unassigned',item_code:r.item_code,item_name:r.item_name,brand:r.brand,qty:0,amount:0};
    agg[k].qty+=r.qty; agg[k].amount+=r.amount;
  });
  const sorted=Object.values(agg).sort((a,b)=>a.sp.localeCompare(b.sp)||b.amount-a.amount);
  const ws=wb.addWorksheet('👤 SP Products',{properties:{tabColor:{argb:'FF'+C.teal}}});
  addSheetHeader(ws,'👤 SP × Product Revenue','Products Sold by Each Salesperson',6);
  setColWidths(ws,[28,18,30,16,14,18]);
  const hRow=ws.getRow(4);
  ['Salesperson','Item Code','Item Name','Brand','Qty','Revenue'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.teal);});
  hRow.height=20; freezeRow(ws,4);
  sorted.forEach((r,i)=>{
    const row=ws.getRow(5+i); const bg=i%2===0?C.white:C.altRow;
    row.getCell(1).value=r.sp;        row.getCell(1).style=cellStyle(bg,true);
    row.getCell(2).value=r.item_code; row.getCell(2).style=cellStyle(bg);
    row.getCell(3).value=r.item_name; row.getCell(3).style=cellStyle(bg);
    row.getCell(4).value=r.brand;     row.getCell(4).style=cellStyle(bg);
    row.getCell(5).value=r.qty;       row.getCell(5).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(6).value=r.amount;    row.getCell(6).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.height=18;
  });
  if(!sorted.length){ws.getRow(5).getCell(1).value='No data.';}
}

function erpBuildRawData(wb, rows) {
  const ws=wb.addWorksheet('📄 Raw Data',{properties:{tabColor:{argb:'FF'+C.midBlue}}});
  addSheetHeader(ws,'📄 Raw Sales Data','All Invoice Line Items — Unfiltered',10);
  setColWidths(ws,[14,16,22,28,14,18,16,14,14,18]);
  const hRow=ws.getRow(4);
  ['Date','Invoice No','Customer','Item Name','Item Code','Brand','Category','Qty','Rate','Amount'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.midBlue);});
  hRow.height=20; freezeRow(ws,4);
  rows.forEach((r,i)=>{
    const row=ws.getRow(5+i); const bg=i%2===0?C.white:C.altRow;
    row.getCell(1).value=r.date;       row.getCell(1).style=cellStyle(bg,false,'center');
    row.getCell(2).value=r.invoice_no; row.getCell(2).style=cellStyle(bg,true);
    row.getCell(3).value=r.customer;   row.getCell(3).style=cellStyle(bg);
    row.getCell(4).value=r.item_name;  row.getCell(4).style=cellStyle(bg);
    row.getCell(5).value=r.item_code;  row.getCell(5).style=cellStyle(bg);
    row.getCell(6).value=r.brand;      row.getCell(6).style=cellStyle(bg);
    row.getCell(7).value=r.category;   row.getCell(7).style=cellStyle(bg);
    row.getCell(8).value=r.qty;        row.getCell(8).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(9).value=r.rate;       row.getCell(9).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(10).value=r.amount;    row.getCell(10).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.height=18;
  });
  if(!rows.length){ws.getRow(5).getCell(1).value='No data.';}
}

function erpBuildSPDrilldown(wb, rows) {
  const bysp={};
  rows.forEach(r=>{
    const k=r.salesperson||'Unassigned';
    if(!bysp[k]) bysp[k]=[];
    bysp[k].push(r);
  });
  const ws=wb.addWorksheet('🔍 SP Drill-down',{properties:{tabColor:{argb:'FF'+C.teal}}});
  addSheetHeader(ws,'🔍 Salesperson Drill-down','Full Transaction List per Salesperson',8);
  setColWidths(ws,[28,14,16,22,14,14,14,18]);
  const hRow=ws.getRow(4);
  ['Salesperson','Date','Invoice No','Customer','Item Code','Brand','Qty','Amount'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.teal);});
  hRow.height=20; freezeRow(ws,4);
  let rowIdx=5;
  Object.keys(bysp).sort().forEach(sp=>{
    bysp[sp].forEach((r,i)=>{
      const row=ws.getRow(rowIdx++); const bg=i%2===0?C.white:C.altRow;
      row.getCell(1).value=r.salesperson||'Unassigned'; row.getCell(1).style=cellStyle(bg,true);
      row.getCell(2).value=r.date;       row.getCell(2).style=cellStyle(bg,false,'center');
      row.getCell(3).value=r.invoice_no; row.getCell(3).style=cellStyle(bg);
      row.getCell(4).value=r.customer;   row.getCell(4).style=cellStyle(bg);
      row.getCell(5).value=r.item_code;  row.getCell(5).style=cellStyle(bg);
      row.getCell(6).value=r.brand;      row.getCell(6).style=cellStyle(bg);
      row.getCell(7).value=r.qty;        row.getCell(7).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
      row.getCell(8).value=r.amount;     row.getCell(8).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
      row.height=18;
    });
  });
  if(!rows.length){ws.getRow(5).getCell(1).value='No data.';}
}

// ── ERP STOCK — DYNAMIC SHEETS (mirrors generate_stock_report.js) ──

function erpBuildStockSummary(wb, data) {
  const stats=data.stats||{};
  const ws=wb.addWorksheet('📊 Summary',{properties:{tabColor:{argb:'FF'+C.darkBlue}}});
  ws.views=[{showGridLines:false}]; setColWidths(ws,[3,36,26,26,3]);
  ws.mergeCells('B2:D2');
  const t=ws.getCell('B2'); t.value='ERPNEXT — STOCK REPORT';
  t.style={font:{bold:true,size:22,name:'Calibri',color:{argb:'FF'+C.headerFont}},fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF'+C.darkBlue}},alignment:{horizontal:'center',vertical:'middle'}};
  ws.getRow(2).height=48;
  ws.mergeCells('B3:D3');
  const s=ws.getCell('B3'); s.value=`Generated: ${fmtDate()}  |  Developed by Arun Chiramal — Full Stack Developer`;
  s.style={font:{bold:false,size:10,name:'Calibri',color:{argb:'FF'+C.headerFont}},fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF'+C.midBlue}},alignment:{horizontal:'center',vertical:'middle'}};
  ws.getRow(3).height=22; ws.getRow(4).height=12;
  const kpis=[
    ['📦 Total Items',       (stats.totalItems||0).toString(),  C.midBlue],
    ['🏷️ Brands',            (stats.brands||0).toString(),      C.purple],
    ['📁 Categories',        (stats.categories||0).toString(),  C.teal],
    ['💰 Total Stock Value', (stats.totalValue||0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}), C.darkBlue],
    ['📊 Total Balance Qty', (stats.totalQty||0).toLocaleString(), C.orange],
    ['⚠️ Zero Stock Items',  (stats.zeroStock||0).toString(),   C.red],
  ];
  kpis.forEach(([label,value,bg],i)=>{
    const row=5+i; ws.mergeCells(`B${row}:C${row}`);
    const lC=ws.getCell(`B${row}`); lC.value=label;
    lC.style={font:{bold:true,size:11,name:'Calibri',color:{argb:'FF333333'}},fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFE9F0FB'}},alignment:{horizontal:'left',vertical:'middle',indent:1},border:thinBorder()};
    const vC=ws.getCell(`D${row}`); vC.value=value;
    vC.style={font:{bold:true,size:14,name:'Calibri',color:{argb:'FF'+C.headerFont}},fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF'+bg}},alignment:{horizontal:'center',vertical:'middle'},border:thinBorder()};
    ws.getRow(row).height=30;
  });
  // Brand quick-view
  ws.getRow(12).height=12;
  ws.mergeCells('B13:D13');
  const bh=ws.getCell('B13'); bh.value='Brand Quick-view';
  bh.style={font:{bold:true,size:11,name:'Calibri',color:{argb:'FF'+C.headerFont}},fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF'+C.midBlue}},alignment:{horizontal:'center',vertical:'middle'}};
  ws.getRow(13).height=22;
  const bHdr=ws.getRow(14); bHdr.height=18;
  ['Brand','Items','Stock Value'].forEach((v,i)=>{
    const c=ws.getCell(`${['B','C','D'][i]}14`); c.value=v; c.style=headerStyle(C.midBlue);
  });
  (stats.brandRows||[]).slice(0,15).forEach((br,i)=>{
    const r=15+i; const bg=i%2===0?C.white:C.altRow;
    ws.getCell(`B${r}`).value=br.brand; ws.getCell(`B${r}`).style=cellStyle(bg,true);
    ws.getCell(`C${r}`).value=br.items; ws.getCell(`C${r}`).style=cellStyle(bg,false,'center');
    ws.getCell(`D${r}`).value=br.value; ws.getCell(`D${r}`).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    ws.getRow(r).height=18;
  });
}

function erpBuildAllStock(wb, allRows) {
  const ws=wb.addWorksheet('📦 All Stock',{properties:{tabColor:{argb:'FF'+C.midBlue}}});
  addSheetHeader(ws,'📦 All Stock','Complete Stock Balance — All Items & Warehouses',12);
  setColWidths(ws,[16,18,32,20,10,12,12,12,14,12,16,14]);
  const hRow=ws.getRow(4);
  ['Brand','Item Code','Item Name','Warehouse','UOM','Opening','In','Out','Balance','Val Rate','Stock Value','Reserved'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.midBlue);});
  hRow.height=20; freezeRow(ws,4);
  allRows.forEach((r,i)=>{
    const row=ws.getRow(5+i); const bg=i%2===0?C.white:C.altRow;
    const zBg=r.Balance_Qty<=0?'FFE4E4':bg;
    row.getCell(1).value=r.Brand;          row.getCell(1).style=cellStyle(bg,true);
    row.getCell(2).value=r.Item_Code;      row.getCell(2).style=cellStyle(bg,true);
    row.getCell(3).value=r.Item_Name;      row.getCell(3).style=cellStyle(bg);
    row.getCell(4).value=r.Warehouse;      row.getCell(4).style=cellStyle(bg);
    row.getCell(5).value=r.UOM;            row.getCell(5).style=cellStyle(bg,false,'center');
    row.getCell(6).value=r.Opening_Qty;    row.getCell(6).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(7).value=r.In_Qty;         row.getCell(7).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(8).value=r.Out_Qty;        row.getCell(8).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(9).value=r.Balance_Qty;    row.getCell(9).style={...cellStyle(zBg,true,'right'),numFmt:'#,##0.00'};
    row.getCell(10).value=r.Val_Rate;      row.getCell(10).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(11).value=r.Balance_Value; row.getCell(11).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(12).value=r.Reserved_Stock;row.getCell(12).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.height=18;
  });
  if(!allRows.length){ws.getRow(5).getCell(1).value='No stock data found.';}
}

function erpBuildBrandSummary(wb, brandRows) {
  const ws=wb.addWorksheet('🏷️ By Brand',{properties:{tabColor:{argb:'FF'+C.purple}}});
  addSheetHeader(ws,'🏷️ Stock by Brand','Totals per Brand',5);
  setColWidths(ws,[28,12,16,18,12]);
  const hRow=ws.getRow(4);
  ['Brand','Items','Total Qty','Stock Value','Rank'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.purple);});
  hRow.height=20; freezeRow(ws,4);
  const sorted=[...brandRows].sort((a,b)=>b.value-a.value);
  sorted.forEach((b,i)=>{
    const row=ws.getRow(5+i); const bg=i%2===0?C.white:C.altRow;
    row.getCell(1).value=b.brand; row.getCell(1).style=cellStyle(bg,true);
    row.getCell(2).value=b.items; row.getCell(2).style=cellStyle(bg,false,'center');
    row.getCell(3).value=b.qty;   row.getCell(3).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(4).value=b.value; row.getCell(4).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(5).value=i+1;     row.getCell(5).style=cellStyle(bg,false,'center');
    row.height=18;
  });
  if(!sorted.length){ws.getRow(5).getCell(1).value='No data.';}
}

function erpBuildCatSummary(wb, catRows) {
  const ws=wb.addWorksheet('📁 By Category',{properties:{tabColor:{argb:'FF'+C.orange}}});
  addSheetHeader(ws,'📁 Stock by Category','Totals per Item Group',4);
  setColWidths(ws,[32,12,16,18]);
  const hRow=ws.getRow(4);
  ['Category','Items','Total Qty','Stock Value'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.orange);});
  hRow.height=20; freezeRow(ws,4);
  const sorted=[...catRows].sort((a,b)=>b.value-a.value);
  sorted.forEach((c,i)=>{
    const row=ws.getRow(5+i); const bg=i%2===0?C.white:C.altRow;
    row.getCell(1).value=c.cat;   row.getCell(1).style=cellStyle(bg,true);
    row.getCell(2).value=c.items; row.getCell(2).style=cellStyle(bg,false,'center');
    row.getCell(3).value=c.qty;   row.getCell(3).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(4).value=c.value; row.getCell(4).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.height=18;
  });
  if(!sorted.length){ws.getRow(5).getCell(1).value='No data.';}
}

function erpBuildBrandSheet(wb, brand, rows, colorHex) {
  const sheetName=`🏷️ ${brand}`.substring(0,30).replace(/[\\/*?[\]:]/g,'');
  const ws=wb.addWorksheet(sheetName,{properties:{tabColor:{argb:colorHex}}});
  addSheetHeader(ws,`🏷️ ${brand}`,`Stock Balance — ${brand} Items`,11);
  setColWidths(ws,[18,32,20,10,12,12,12,14,12,16,14]);
  const hRow=ws.getRow(4);
  ['Item Code','Item Name','Warehouse','UOM','Opening','In','Out','Balance','Val Rate','Stock Value','Reserved'].forEach((h,i)=>{
    hRow.getCell(i+1).value=h;
    hRow.getCell(i+1).style={...headerStyle('FFFFFF'),fill:{type:'pattern',pattern:'solid',fgColor:{argb:colorHex}}};
  });
  hRow.height=20; freezeRow(ws,4);
  rows.forEach((r,i)=>{
    const row=ws.getRow(5+i); const bg=i%2===0?C.white:C.altRow;
    const zBg=r.Balance_Qty<=0?'FFE4E4':bg;
    row.getCell(1).value=r.Item_Code;      row.getCell(1).style=cellStyle(bg,true);
    row.getCell(2).value=r.Item_Name;      row.getCell(2).style=cellStyle(bg);
    row.getCell(3).value=r.Warehouse;      row.getCell(3).style=cellStyle(bg);
    row.getCell(4).value=r.UOM;            row.getCell(4).style=cellStyle(bg,false,'center');
    row.getCell(5).value=r.Opening_Qty;    row.getCell(5).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(6).value=r.In_Qty;         row.getCell(6).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(7).value=r.Out_Qty;        row.getCell(7).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(8).value=r.Balance_Qty;    row.getCell(8).style={...cellStyle(zBg,true,'right'),numFmt:'#,##0.00'};
    row.getCell(9).value=r.Val_Rate;       row.getCell(9).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(10).value=r.Balance_Value; row.getCell(10).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(11).value=r.Reserved_Stock;row.getCell(11).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.height=18;
  });
  const totRow=ws.getRow(5+rows.length); totRow.height=22;
  const totals=[null,null,null,null,
    rows.reduce((s,r)=>s+r.Opening_Qty,0),rows.reduce((s,r)=>s+r.In_Qty,0),rows.reduce((s,r)=>s+r.Out_Qty,0),
    rows.reduce((s,r)=>s+r.Balance_Qty,0),null,rows.reduce((s,r)=>s+r.Balance_Value,0),rows.reduce((s,r)=>s+r.Reserved_Stock,0),
  ];
  totals.forEach((v,i)=>{
    const c=totRow.getCell(i+1);
    c.style={...headerStyle('FFFFFF'),fill:{type:'pattern',pattern:'solid',fgColor:{argb:colorHex}},numFmt:'#,##0.00'};
    if(i===0){c.value='TOTAL';} else if(v!==null){c.value=v;}
  });
  if(!rows.length){ws.getRow(5).getCell(1).value='No items for this brand.';}
}

// ── ERP AUDIT — UP TO 7 SHEETS (mirrors audit_sales_data.js) ─

function erpBuildAuditSummary(wb, data) {
  const f=data.findings||{}; const summary=f.summary||[];
  const ws=wb.addWorksheet('🚨 Audit Summary',{properties:{tabColor:{argb:'FF'+C.red}}});
  ws.views=[{showGridLines:false}]; setColWidths(ws,[3,34,10,40,3]);
  ws.mergeCells('B2:D2');
  const t=ws.getCell('B2'); t.value='ERPNEXT — SALES AUDIT REPORT';
  t.style={font:{bold:true,size:22,name:'Calibri',color:{argb:'FF'+C.headerFont}},fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF'+C.red}},alignment:{horizontal:'center',vertical:'middle'}};
  ws.getRow(2).height=48;
  ws.mergeCells('B3:D3');
  const s=ws.getCell('B3');
  const auditPeriod = data.periodYear ? `01 Jan ${data.periodYear} – 31 Dec ${data.periodYear}  |  ` : '';
  s.value=`${auditPeriod}Generated: ${fmtDate()}  |  Invoices: ${(data.invoices||[]).length}  |  Total Revenue: ${(data.totalRev||0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})}`;
  s.style={font:{bold:false,size:10,name:'Calibri',color:{argb:'FF'+C.headerFont}},fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFAA0000'}},alignment:{horizontal:'center',vertical:'middle'}};
  ws.getRow(3).height=22; ws.getRow(4).height=12;
  const hRow=ws.getRow(5); hRow.height=22;
  ['B','C','D'].forEach((col,i)=>{
    const c=ws.getCell(`${col}5`); c.value=['Audit Check','Count','Finding'][i]; c.style=headerStyle(C.red);
  });
  summary.forEach((item,i)=>{
    const row=6+i; ws.getRow(row).height=24;
    const bg=item.risk==='HIGH'?'FFC7CE':item.risk==='MEDIUM'?'FFEB9C':item.risk==='CLEAR'?'C6EFCE':'D9E1F2';
    const fg=item.risk==='HIGH'?'FF9C0006':item.risk==='MEDIUM'?'FF7D6608':item.risk==='CLEAR'?'FF375623':'FF1F3864';
    ws.getCell(`B${row}`).value=item.category; ws.getCell(`B${row}`).style={...cellStyle(bg,true),font:{bold:true,size:11,name:'Calibri',color:{argb:fg}}};
    ws.getCell(`C${row}`).value=item.count;    ws.getCell(`C${row}`).style={...cellStyle(bg,true,'center'),font:{bold:true,size:12,name:'Calibri',color:{argb:fg}}};
    ws.getCell(`D${row}`).value=item.detail;   ws.getCell(`D${row}`).style={...cellStyle(bg,false),alignment:{vertical:'middle',wrapText:true}};
  });
  if(!summary.length){ws.getCell('B6').value='No findings computed.';}
}

function erpBuildTestItems(wb, findings) {
  const items=findings.testItems||[];
  const ws=wb.addWorksheet('🔴 Test Items',{properties:{tabColor:{argb:'FFC00000'}}});
  addSheetHeader(ws,'🔴 Test Items in Live Invoices','TEST items found in submitted sales invoices',7);
  setColWidths(ws,[10,14,22,28,18,16,18]);
  const hRow=ws.getRow(4);
  ['Risk','Invoice','Customer','Item Name','Item Code','Date','Amount'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.red);});
  hRow.height=20; freezeRow(ws,4);
  items.forEach((f,i)=>{
    const row=ws.getRow(5+i); const bg=f.risk==='HIGH'?'FFC7CE':'FCE4E4';
    row.getCell(1).value=f.risk;      row.getCell(1).style=cellStyle(bg,true,'center');
    row.getCell(2).value=f.invoice;   row.getCell(2).style=cellStyle(bg);
    row.getCell(3).value=f.customer;  row.getCell(3).style=cellStyle(bg);
    row.getCell(4).value=f.item_name; row.getCell(4).style=cellStyle(bg);
    row.getCell(5).value=f.item_code; row.getCell(5).style=cellStyle(bg);
    row.getCell(6).value=f.date;      row.getCell(6).style=cellStyle(bg,false,'center');
    row.getCell(7).value=f.amount;    row.getCell(7).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.height=18;
  });
}

function erpBuildNegativeAmounts(wb, findings) {
  const items=findings.negativeAmounts||[];
  const ws=wb.addWorksheet('🔴 Negative Amounts',{properties:{tabColor:{argb:'FFC00000'}}});
  addSheetHeader(ws,'🔴 Negative Amount Line Items','Possible errors or unauthorised credit notes',7);
  setColWidths(ws,[10,14,22,28,18,16,18]);
  const hRow=ws.getRow(4);
  ['Risk','Invoice','Customer','Item Name','Item Code','Date','Amount'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.red);});
  hRow.height=20; freezeRow(ws,4);
  items.forEach((f,i)=>{
    const row=ws.getRow(5+i); const bg=f.risk==='HIGH'?'FFC7CE':'FFEB9C';
    row.getCell(1).value=f.risk;      row.getCell(1).style=cellStyle(bg,true,'center');
    row.getCell(2).value=f.invoice;   row.getCell(2).style=cellStyle(bg);
    row.getCell(3).value=f.customer;  row.getCell(3).style=cellStyle(bg);
    row.getCell(4).value=f.item_name; row.getCell(4).style=cellStyle(bg);
    row.getCell(5).value=f.item_code; row.getCell(5).style=cellStyle(bg);
    row.getCell(6).value=f.date;      row.getCell(6).style=cellStyle(bg,false,'center');
    row.getCell(7).value=f.amount;    row.getCell(7).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.height=18;
  });
}

function erpBuildPriceVariance(wb, findings) {
  const items=findings.priceVariance||[];
  const ws=wb.addWorksheet('🟡 Price Variance',{properties:{tabColor:{argb:'FFFFC000'}}});
  addSheetHeader(ws,'🟡 Price Inconsistency','Same item sold at significantly different prices to different customers',8);
  setColWidths(ws,[10,18,28,10,14,14,14,28]);
  const hRow=ws.getRow(4);
  ['Risk','Item Code','Item Name','Times Sold','Min Rate','Max Rate','Variance %','Note'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle('FFC000');});
  hRow.height=20; freezeRow(ws,4);
  items.forEach((f,i)=>{
    const row=ws.getRow(5+i); const bg=f.risk==='HIGH'?'FFC7CE':f.risk==='MEDIUM'?'FFEB9C':'F0F0F0';
    row.getCell(1).value=f.risk;         row.getCell(1).style=cellStyle(bg,true,'center');
    row.getCell(2).value=f.item_code;    row.getCell(2).style=cellStyle(bg);
    row.getCell(3).value=f.item_name;    row.getCell(3).style=cellStyle(bg);
    row.getCell(4).value=f.times_sold;   row.getCell(4).style=cellStyle(bg,false,'center');
    row.getCell(5).value=f.min_rate;     row.getCell(5).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(6).value=f.max_rate;     row.getCell(6).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(7).value=f.variance_pct; row.getCell(7).style={...cellStyle(bg,true,'right'),numFmt:'0"%"'};
    row.getCell(8).value=f.note;         row.getCell(8).style={...cellStyle(bg),alignment:{vertical:'middle',wrapText:true}};
    row.height=24;
  });
}

function erpBuildUnassignedSP(wb, findings) {
  const items=findings.unassignedSP||[];
  const ws=wb.addWorksheet('🟡 No Salesperson',{properties:{tabColor:{argb:'FFFFC000'}}});
  addSheetHeader(ws,'🟡 Revenue Without Salesperson','Invoices with no salesperson assigned',6);
  setColWidths(ws,[10,22,26,14,14,18]);
  const hRow=ws.getRow(4);
  ['Risk','Invoice','Customer','Date','Line Items','Revenue'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle('FBBF24');});
  hRow.height=20; freezeRow(ws,4);
  items.forEach((f,i)=>{
    const row=ws.getRow(5+i); const bg=f.risk==='MEDIUM'?'FFEB9C':'F9F9F9';
    row.getCell(1).value=f.risk;       row.getCell(1).style=cellStyle(bg,true,'center');
    row.getCell(2).value=f.invoice;    row.getCell(2).style=cellStyle(bg);
    row.getCell(3).value=f.customer;   row.getCell(3).style=cellStyle(bg);
    row.getCell(4).value=f.date;       row.getCell(4).style=cellStyle(bg,false,'center');
    row.getCell(5).value=f.line_items; row.getCell(5).style=cellStyle(bg,false,'center');
    row.getCell(6).value=f.revenue;    row.getCell(6).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.height=18;
  });
}

function erpBuildWeekendInvoices(wb, findings) {
  const items=findings.weekendInvoices||[];
  const ws=wb.addWorksheet('🟡 Weekend Invoices',{properties:{tabColor:{argb:'FFFFC000'}}});
  addSheetHeader(ws,'🟡 Weekend Invoices','Invoices dated on Friday or Saturday (Saudi weekend)',6);
  setColWidths(ws,[10,28,14,14,12,18]);
  const hRow=ws.getRow(4);
  ['Risk','Invoice(s)','Date','Day','Count','Revenue'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle('FBBF24');});
  hRow.height=20; freezeRow(ws,4);
  items.forEach((f,i)=>{
    const row=ws.getRow(5+i); const bg=f.risk==='MEDIUM'?'FFEB9C':'F9F9F9';
    row.getCell(1).value=f.risk;      row.getCell(1).style=cellStyle(bg,true,'center');
    row.getCell(2).value=f.invoice;   row.getCell(2).style=cellStyle(bg);
    row.getCell(3).value=f.date;      row.getCell(3).style=cellStyle(bg,false,'center');
    row.getCell(4).value=f.day;       row.getCell(4).style=cellStyle(bg,false,'center');
    row.getCell(5).value=f.inv_count; row.getCell(5).style=cellStyle(bg,false,'center');
    row.getCell(6).value=f.revenue;   row.getCell(6).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.height=18;
  });
}

function erpBuildConcentration(wb, findings) {
  const items=findings.concentration||[];
  const ws=wb.addWorksheet('🔵 Concentration',{properties:{tabColor:{argb:'FF'+C.midBlue}}});
  addSheetHeader(ws,'🔵 Customer Concentration','Top 10 Customers by Revenue Share',5);
  setColWidths(ws,[10,36,18,14,14]);
  const hRow=ws.getRow(4);
  ['Risk','Customer','Revenue','% of Total','Rank'].forEach((h,i)=>{hRow.getCell(i+1).value=h;hRow.getCell(i+1).style=headerStyle(C.midBlue);});
  hRow.height=20; freezeRow(ws,4);
  items.forEach((f,i)=>{
    const row=ws.getRow(5+i); const bg=f.risk==='HIGH'?'FFC7CE':f.risk==='MEDIUM'?'FFEB9C':'D9E1F2';
    row.getCell(1).value=f.risk;     row.getCell(1).style=cellStyle(bg,true,'center');
    row.getCell(2).value=f.customer; row.getCell(2).style=cellStyle(bg,true);
    row.getCell(3).value=f.revenue;  row.getCell(3).style={...cellStyle(bg,false,'right'),numFmt:'#,##0.00'};
    row.getCell(4).value=f.pct;      row.getCell(4).style={...cellStyle(bg,true,'right'),numFmt:'0.0"%"'};
    row.getCell(5).value=i+1;        row.getCell(5).style=cellStyle(bg,false,'center');
    row.height=18;
  });
  if(!items.length){ws.getRow(5).getCell(1).value='No customer data.';}
}

// ── RAW DATA SHEET ────────────────────────────────────────────
// Dumps the unprocessed Odoo API response so users can see every field exactly as Odoo returns it.
function buildOdooRawData(wb, rawLeads, rawContacts) {
  // ── CRM Raw Data ──────────────────────────────────────────────
  const ws = wb.addWorksheet('🗃️ Raw Data', { properties: { tabColor: { argb: 'FF4B5563' } } });
  ws.views = [{ showGridLines: true, state: 'frozen', ySplit: 2 }];

  addSheetHeader(ws, '🗃️ Raw Data — CRM Leads & Contacts', 'Unprocessed data exactly as returned by Odoo API', 14);

  // CRM Leads section
  const leadFields = ['id','name','partner_name','email_from','phone','planned_revenue','stage_id','probability','date_deadline','user_id','kanban_state','type','create_date'];
  const leadHeaders = ['ID','Deal / Lead Name','Company','Email','Phone','Value','Stage','Probability %','Close Date','Owner','Kanban State','Type','Created Date'];

  const hRow = ws.addRow(leadHeaders);
  hRow.height = 24;
  hRow.eachCell(cell => {
    cell.font      = { bold: true, color: { argb: 'FF' + C.headerFont }, size: 11 };
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.darkBlue } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
    cell.border    = { bottom: { style: 'thin', color: { argb: 'FF' + C.borderColor } } };
  });

  (rawLeads || []).forEach((rec, i) => {
    const row = ws.addRow(leadFields.map(f => {
      const v = rec[f];
      if (Array.isArray(v)) return v[1] || v[0] || '';  // Odoo many2one returns [id, name]
      return v !== undefined && v !== false ? v : '';
    }));
    if (i % 2 === 1) {
      row.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.altRow } };
      });
    }
    row.eachCell(cell => {
      cell.alignment = { vertical: 'middle' };
      cell.border    = { bottom: { style: 'hair', color: { argb: 'FFD1D5DB' } } };
    });
  });

  if (!rawLeads || rawLeads.length === 0) {
    ws.addRow(['No CRM lead/opportunity records returned by Odoo.']);
  }

  // Column widths
  const colWidths = [8, 36, 28, 28, 18, 14, 20, 14, 14, 22, 16, 14, 18];
  leadFields.forEach((_, i) => { ws.getColumn(i + 1).width = colWidths[i] || 16; });

  // Spacer + Contacts section header
  ws.addRow([]);
  const cLabel = ws.addRow(['CONTACTS — Raw Data']);
  cLabel.getCell(1).font = { bold: true, size: 12, color: { argb: 'FF' + C.darkBlue } };

  const contactFields = ['id','name','email','phone','mobile','company_name','job_position','create_date'];
  const contactHeaders = ['ID','Full Name','Email','Phone','Mobile','Company','Job Title','Created Date'];
  const chRow = ws.addRow(contactHeaders);
  chRow.height = 24;
  chRow.eachCell(cell => {
    cell.font      = { bold: true, color: { argb: 'FF' + C.headerFont }, size: 11 };
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.midBlue } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border    = { bottom: { style: 'thin', color: { argb: 'FF' + C.borderColor } } };
  });

  (rawContacts || []).forEach((rec, i) => {
    const row = ws.addRow(contactFields.map(f => {
      const v = rec[f];
      if (Array.isArray(v)) return v[1] || v[0] || '';
      return v !== undefined && v !== false ? v : '';
    }));
    if (i % 2 === 1) {
      row.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.altRow } };
      });
    }
    row.eachCell(cell => {
      cell.alignment = { vertical: 'middle' };
      cell.border    = { bottom: { style: 'hair', color: { argb: 'FFD1D5DB' } } };
    });
  });

  if (!rawContacts || rawContacts.length === 0) {
    ws.addRow(['No contact records returned by Odoo.']);
  }
}

// ── LEADS & CONTACTS DASHBOARD ────────────────────────────────
function buildLeadsContactsDashboard(wb, data) {
  const ws = wb.addWorksheet('📊 Dashboard', { properties: { tabColor: { argb: 'FF' + C.indigo } } });
  ws.views = [{ showGridLines: false }];
  setColWidths(ws, [4, 28, 22, 22, 22, 22, 4]);
  addSheetHeader(ws, '📊 Leads & Contacts Overview', toStdDate(new Date()), 6);
  const leads = data.leads || []; const contacts = data.contacts || [];
  const openLeads = leads.filter(l => !['won','lost','converted'].includes((l.status||'').toLowerCase())).length;
  const stats = [
    ['Total Leads', leads.length], ['Open Leads', openLeads],
    ['Converted', leads.filter(l => (l.status||'').toLowerCase() === 'converted').length],
    ['Total Contacts', contacts.length],
    ['With Email', contacts.filter(c => c.email).length],
    ['With Phone', contacts.filter(c => c.phone).length],
  ];
  let r = 5;
  ws.getRow(r).values = ['', 'Metric', '', 'Count', '', '']; applyHeader(ws.getRow(r), 6); r++;
  stats.forEach(([label, val]) => {
    ws.getRow(r).values = ['', label, '', val, '', ''];
    ws.getCell(`B${r}`).style = cellStyle(C.rowAlt); ws.getCell(`D${r}`).style = { ...cellStyle(), font: { bold: true, size: 13 } };
    r++;
  });
}

// ── LEAD SOURCE ANALYSIS ─────────────────────────────────────
function buildLeadSourceAnalysis(wb, leads) {
  const ws = wb.addWorksheet('📡 Lead Source', { properties: { tabColor: { argb: 'FF' + C.orange } } });
  ws.views = [{ showGridLines: false }];
  setColWidths(ws, [4, 28, 16, 16, 22, 4]);
  addSheetHeader(ws, '📡 Lead Source Analysis', 'Where your leads are coming from', 5);
  const sourceMap = {};
  (leads || []).forEach(l => {
    const src = l.source || 'Unknown';
    if (!sourceMap[src]) sourceMap[src] = { total: 0, open: 0, converted: 0 };
    sourceMap[src].total++;
    if ((l.status||'').toLowerCase() === 'converted') sourceMap[src].converted++;
    else sourceMap[src].open++;
  });
  const rows = Object.entries(sourceMap).sort((a, b) => b[1].total - a[1].total);
  let r = 5;
  ws.getRow(r).values = ['', 'Source', 'Total Leads', 'Converted', 'Conversion Rate', '']; applyHeader(ws.getRow(r), 5); r++;
  rows.forEach(([src, d], i) => {
    const rate = d.total > 0 ? ((d.converted / d.total) * 100).toFixed(1) + '%' : '0%';
    ws.getRow(r).values = ['', src, d.total, d.converted, rate, ''];
    const bg = i % 2 === 0 ? 'FFFFFFFF' : C.rowAlt;
    [2,3,4,5].forEach(c => ws.getCell(r, c).style = cellStyle(bg));
    r++;
  });
  // Totals
  const total = (leads||[]).length;
  const converted = (leads||[]).filter(l => (l.status||'').toLowerCase() === 'converted').length;
  ws.getRow(r).values = ['', 'TOTAL', total, converted, total > 0 ? ((converted/total)*100).toFixed(1)+'%' : '0%', ''];
  applyHeader(ws.getRow(r), 5); freezeRow(ws, 5);
}

// ── ACTIVITIES DASHBOARD ─────────────────────────────────────
function buildActivitiesDashboard(wb, data) {
  const ws = wb.addWorksheet('📊 Dashboard', { properties: { tabColor: { argb: 'FF' + C.indigo } } });
  ws.views = [{ showGridLines: false }];
  setColWidths(ws, [4, 28, 22, 22, 22, 4]);
  addSheetHeader(ws, '📊 Activity Overview', toStdDate(new Date()), 5);
  const acts = data.activities || [];
  const now = new Date();
  const overdue = acts.filter(a => a.dueDate && new Date(a.dueDate) < now && (a.status||'').toLowerCase() !== 'done').length;
  const byType = {};
  acts.forEach(a => { const t = a.type || 'Other'; byType[t] = (byType[t] || 0) + 1; });
  const stats = [['Total Activities', acts.length], ['Overdue', overdue], ['Done', acts.filter(a => (a.status||'').toLowerCase() === 'done').length], ...Object.entries(byType)];
  let r = 5;
  ws.getRow(r).values = ['', 'Metric', '', 'Count', '', '']; applyHeader(ws.getRow(r), 5); r++;
  stats.forEach(([label, val]) => {
    ws.getRow(r).values = ['', label, '', val, '', ''];
    ws.getCell(`B${r}`).style = cellStyle(C.rowAlt);
    ws.getCell(`D${r}`).style = { ...cellStyle(), font: { bold: true, size: 13 } };
    r++;
  });
}

// ── OVERDUE ACTIVITIES ────────────────────────────────────────
function buildOverdueActivities(wb, activities) {
  const ws = wb.addWorksheet('⚠️ Overdue', { properties: { tabColor: { argb: 'FFDC2626' } } });
  ws.views = [{ showGridLines: false }];
  setColWidths(ws, [4, 28, 18, 18, 20, 20, 4]);
  addSheetHeader(ws, '⚠️ Overdue Activities', 'Tasks past their due date', 6);
  const now = new Date();
  const overdue = (activities || []).filter(a => a.dueDate && new Date(a.dueDate) < now && (a.status||'').toLowerCase() !== 'done')
    .sort((a, b) => new Date(a.dueDate) - new Date(b.dueDate));
  let r = 5;
  ws.getRow(r).values = ['', 'Activity', 'Type', 'Due Date', 'Assigned To', 'Related To', '']; applyHeader(ws.getRow(r), 6); r++;
  overdue.forEach((a, i) => {
    ws.getRow(r).values = ['', a.summary || a.description || '—', a.type || '—', toStdDate(a.dueDate), a.user || '—', a.relatedTo || '—', ''];
    const bg = i % 2 === 0 ? 'FFFEF2F2' : 'FFFFFFFF';
    [2,3,4,5,6].forEach(c => ws.getCell(r, c).style = cellStyle(bg));
    r++;
  });
  if (overdue.length === 0) { ws.getRow(r).getCell(2).value = '✅ No overdue activities — all up to date!'; }
  freezeRow(ws, 5);
}

// ── ACTIVITIES BY USER ────────────────────────────────────────
function buildActivitiesByUser(wb, activities) {
  const ws = wb.addWorksheet('👤 By User', { properties: { tabColor: { argb: 'FF' + C.purple } } });
  ws.views = [{ showGridLines: false }];
  setColWidths(ws, [4, 28, 14, 14, 14, 18, 4]);
  addSheetHeader(ws, '👤 Activity by User', 'Workload distribution across team members', 6);
  const userMap = {};
  (activities || []).forEach(a => {
    const u = a.user || 'Unassigned';
    if (!userMap[u]) userMap[u] = { total: 0, done: 0, overdue: 0, types: {} };
    userMap[u].total++;
    if ((a.status||'').toLowerCase() === 'done') userMap[u].done++;
    if (a.dueDate && new Date(a.dueDate) < new Date() && (a.status||'').toLowerCase() !== 'done') userMap[u].overdue++;
    const t = a.type || 'Other'; userMap[u].types[t] = (userMap[u].types[t] || 0) + 1;
  });
  const rows = Object.entries(userMap).sort((a, b) => b[1].total - a[1].total);
  const medals = ['🥇', '🥈', '🥉'];
  let r = 5;
  ws.getRow(r).values = ['', 'User', 'Total', 'Done', 'Overdue', 'Completion Rate', '']; applyHeader(ws.getRow(r), 6); r++;
  rows.forEach(([user, d], i) => {
    const rate = d.total > 0 ? ((d.done / d.total) * 100).toFixed(1) + '%' : '0%';
    const medal = medals[i] || '';
    ws.getRow(r).values = ['', `${medal} ${user}`, d.total, d.done, d.overdue, rate, ''];
    const bg = i % 2 === 0 ? 'FFFFFFFF' : C.rowAlt;
    [2,3,4,5,6].forEach(c => ws.getCell(r, c).style = cellStyle(bg));
    if (d.overdue > 0) ws.getCell(r, 5).style = { ...cellStyle(bg), font: { color: { argb: 'FFDC2626' }, bold: true } };
    r++;
  });
  freezeRow(ws, 5);
}

// ── WARNINGS SHEET ────────────────────────────────────────────
function buildWarningsSheet(wb, warnings, crmName) {
  const ws = wb.addWorksheet('⚠️ Warnings', { properties: { tabColor: { argb: 'FFED7D31' } } });
  ws.views = [{ showGridLines: false }];
  setColWidths(ws, [6, 60, 30]);

  ws.mergeCells('A1:C1');
  const t = ws.getCell('A1');
  t.value = `⚠️ ${crmName} — Report Generated with Warnings`;
  t.style = { font:{bold:true,size:14,name:'Calibri',color:{argb:'FFFFFFFF'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFED7D31'}}, alignment:{horizontal:'center',vertical:'middle'} };
  ws.getRow(1).height = 36;

  ws.mergeCells('A2:C2');
  const s = ws.getCell('A2');
  s.value = `Generated: ${fmtDate()}  |  Some data could not be fetched — see details below`;
  s.style = { font:{bold:false,size:10,name:'Calibri',color:{argb:'FFFFFFFF'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFC55A11'}}, alignment:{horizontal:'center',vertical:'middle'} };
  ws.getRow(2).height = 20;

  ws.getRow(3).height = 8;

  const hRow = ws.getRow(4);
  hRow.getCell(1).value = '#';       hRow.getCell(1).style = headerStyle(C.orange);
  hRow.getCell(2).value = 'Error / Warning Message'; hRow.getCell(2).style = headerStyle(C.orange);
  hRow.getCell(3).value = 'What To Do'; hRow.getCell(3).style = headerStyle(C.orange);
  hRow.height = 20;

  const fixes = {
    '401': 'Generate a new API token from your CRM and try again.',
    '403': 'Check the required scopes/permissions in the setup guide.',
    '404': 'Check your CRM URL — endpoint not found.',
    '429': 'Wait 60 seconds then run the report again.',
    'timed out': 'Try a focused sub-report (Pipeline or Contacts) instead of Full Report.',
    'login failed': 'Check your ERPNext username and password.',
    'cannot reach': 'Check your ERPNext URL is correct and accessible.',
  };

  warnings.forEach((w, i) => {
    const row = ws.getRow(5 + i);
    const bg  = i % 2 === 0 ? 'FFFDF3EC' : C.white;
    let fix   = 'Review your credentials and check the setup guide at /docs/setup.html';
    for (const [key, val] of Object.entries(fixes)) {
      if (w.toLowerCase().includes(key)) { fix = val; break; }
    }
    row.getCell(1).value = i + 1;  row.getCell(1).style = cellStyle(bg, true, 'center');
    row.getCell(2).value = w;       row.getCell(2).style = { ...cellStyle('FFFFE8D6'), font:{bold:false,size:10,name:'Calibri',color:{argb:'FF7C2D00'}}, border:thinBorder(), alignment:{vertical:'middle',wrapText:true} };
    row.getCell(3).value = fix;     row.getCell(3).style = cellStyle(bg, false, 'left');
    row.height = 28;
  });

  ws.mergeCells(`A${5+warnings.length+1}:C${5+warnings.length+1}`);
  const note = ws.getCell(`A${5+warnings.length+1}`);
  note.value = '📖 Full setup guides at: https://crm-data-extractor-taupe.vercel.app/docs/setup.html';
  note.style = { font:{italic:true,size:10,name:'Calibri',color:{argb:'FF6B7280'}}, alignment:{horizontal:'center'} };
}

// ── FOCUSED DASHBOARD (for sub-reports) ──────────────────────
function buildFocusedDashboard(wb, data, focus) {
  const ws = wb.addWorksheet('📊 Dashboard', { properties: { tabColor: { argb: 'FF' + C.darkBlue } } });
  ws.views = [{ showGridLines: false }];
  setColWidths(ws, [3, 30, 25, 25, 3]);

  const focusLabel = { pipeline:'PIPELINE REPORT', contacts:'CONTACTS REPORT', leads:'LEADS REPORT', activities:'ACTIVITIES REPORT' }[focus] || 'REPORT';

  ws.mergeCells('B2:D2');
  const title = ws.getCell('B2');
  title.value = `${data.crmName.toUpperCase()} — ${focusLabel}`;
  title.style = { font:{bold:true,size:20,name:'Calibri',color:{argb:'FF'+C.headerFont}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF'+C.darkBlue}}, alignment:{horizontal:'center',vertical:'middle'} };
  ws.getRow(2).height = 44;

  ws.mergeCells('B3:D3');
  const sub = ws.getCell('B3');
  sub.value = `Generated: ${fmtDate()}  |  Developed by Arun Chiramal — Full Stack Developer`;
  sub.style = { font:{bold:false,size:10,name:'Calibri',color:{argb:'FF'+C.headerFont}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF'+C.midBlue}}, alignment:{horizontal:'center',vertical:'middle'} };
  ws.getRow(3).height = 22;
  ws.getRow(4).height = 12;

  const kpis = [];
  if (focus === 'pipeline') {
    kpis.push(['💰 Total Pipeline Value', `$${(data.kpis.totalPipeline||0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})}`, C.darkBlue]);
    kpis.push(['📂 Open Deals',           String(data.kpis.openDeals||0),     C.midBlue]);
    kpis.push(['⏰ Closing in 30 Days',   String(data.kpis.closingSoon||0),   C.orange]);
    kpis.push(['✅ Won This Month',        String(data.kpis.wonThisMonth||0),  '375623']);
  } else if (focus === 'contacts') {
    kpis.push(['👤 Total Contacts',       String(data.contacts.length||0),    C.green]);
    kpis.push(['📋 Records Extracted',    String(data.contacts.length||0),    C.midBlue]);
  } else if (focus === 'leads') {
    kpis.push(['🔍 Total Leads',          String(data.leads.length||0),       C.teal]);
    kpis.push(['📋 Records Extracted',    String(data.leads.length||0),       C.midBlue]);
  } else if (focus === 'activities') {
    kpis.push(['🔔 Total Activities',     String(data.activities.length||0),  C.purple]);
    kpis.push(['📋 Records Extracted',    String(data.activities.length||0),  C.midBlue]);
  }

  kpis.forEach(([label, value, bg], i) => {
    const row = 5 + i;
    ws.mergeCells(`B${row}:C${row}`);
    const lCell = ws.getCell(`B${row}`);
    lCell.value = label;
    lCell.style = { font:{bold:true,size:11,name:'Calibri',color:{argb:'FF333333'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFE9F0FB'}}, alignment:{horizontal:'left',vertical:'middle',indent:1}, border:thinBorder() };
    const vCell = ws.getCell(`D${row}`);
    vCell.value = value;
    vCell.style = { font:{bold:true,size:14,name:'Calibri',color:{argb:'FF'+C.headerFont}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF'+bg}}, alignment:{horizontal:'center',vertical:'middle'}, border:thinBorder() };
    ws.getRow(row).height = 30;
  });
}

// ── ODOO — FORECAST SHEET ────────────────────────────────────
// Weighted pipeline = amount × (probability / 100), grouped by close month
function buildForecastSheet(wb, pipeline) {
  const ws = wb.addWorksheet('🔮 Forecast', { properties: { tabColor: { argb: 'FF002060' } } });
  ws.views = [{ showGridLines: false, state: 'frozen', ySplit: 5 }];
  setColWidths(ws, [16, 14, 20, 20, 16, 16]);

  ws.mergeCells('A2:F2');
  const t = ws.getCell('A2');
  t.value = 'REVENUE FORECAST — WEIGHTED PIPELINE';
  t.style = { font: { bold: true, size: 16, name: 'Calibri', color: { argb: 'FF' + C.headerFont } },
              fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF002060' } },
              alignment: { horizontal: 'center', vertical: 'middle' } };
  ws.getRow(2).height = 40;

  ws.mergeCells('A3:F3');
  const sub = ws.getCell('A3');
  sub.value = 'Weighted Value = Deal Amount × Probability %  |  Excludes Won & Lost deals  |  Generated: ' + fmtDate();
  sub.style = { font: { size: 10, name: 'Calibri', color: { argb: 'FF' + C.headerFont } },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.midBlue } },
                alignment: { horizontal: 'center', vertical: 'middle' } };
  ws.getRow(3).height = 20;
  ws.getRow(4).height = 8;

  const hRow = ws.getRow(5);
  ['Close Month', 'Deal Count', 'Full Pipeline Value', 'Weighted Forecast', 'Avg Probability %', 'Best Case Value'].forEach((h, i) => {
    const c = hRow.getCell(i + 1);
    c.value = h;
    c.style = headerStyle(C.darkBlue);
  });
  ws.getRow(5).height = 24;

  // Only active deals (not won/lost)
  const active = pipeline.filter(d => {
    const s = (d.stage || '').toLowerCase();
    return !s.includes('won') && !s.includes('lost');
  });

  // Group by close month (YYYY-MM from closeDate)
  const monthMap = {};
  active.forEach(d => {
    // closeDate is already DD MMM YYYY — parse back to get month key
    let mk = 'No Date';
    if (d.closeDate) {
      const parsed = new Date(d.closeDate);
      if (!isNaN(parsed)) mk = `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, '0')}`;
    }
    if (!monthMap[mk]) monthMap[mk] = { deals: 0, pipeline: 0, weighted: 0, probSum: 0, bestCase: 0 };
    const prob = d.probability || 0;
    monthMap[mk].deals++;
    monthMap[mk].pipeline  += d.amount || 0;
    monthMap[mk].weighted  += (d.amount || 0) * (prob / 100);
    monthMap[mk].probSum   += prob;
    monthMap[mk].bestCase  += prob >= 50 ? (d.amount || 0) : 0;
  });

  const months = Object.keys(monthMap).sort();
  months.forEach((mk, ri) => {
    const agg = monthMap[mk];
    const avgProb = agg.deals ? (agg.probSum / agg.deals) : 0;
    const r = ws.getRow(6 + ri);
    const bg = ri % 2 === 0 ? C.white : C.altRow;
    // Highlight high-confidence months (avg prob > 60%)
    const highConf = avgProb >= 60;
    [mk, agg.deals, agg.pipeline, agg.weighted, avgProb, agg.bestCase].forEach((v, ci) => {
      const c = r.getCell(ci + 1);
      c.value = v;
      c.style = {
        ...cellStyle(highConf ? 'F0FDF4' : bg, highConf, ci === 0 ? 'left' : 'right'),
        numFmt: ci === 2 || ci === 3 || ci === 5 ? '#,##0.00' : ci === 4 ? '0.0' : 'General',
      };
      if (highConf && ci === 4) c.style.font = { ...c.style.font, bold: true, color: { argb: 'FF' + C.green } };
    });
    r.height = 18;
  });

  // Totals row
  if (months.length) {
    const totRow = ws.getRow(6 + months.length);
    const totDeals    = months.reduce((s, m) => s + monthMap[m].deals, 0);
    const totPipeline = months.reduce((s, m) => s + monthMap[m].pipeline, 0);
    const totWeighted = months.reduce((s, m) => s + monthMap[m].weighted, 0);
    const totBest     = months.reduce((s, m) => s + monthMap[m].bestCase, 0);
    const totAvgProb  = totDeals ? months.reduce((s, m) => s + monthMap[m].probSum, 0) / totDeals : 0;
    ['TOTAL', totDeals, totPipeline, totWeighted, totAvgProb, totBest].forEach((v, ci) => {
      const c = totRow.getCell(ci + 1);
      c.value = v;
      c.style = { ...headerStyle(C.darkBlue), numFmt: ci === 2 || ci === 3 || ci === 5 ? '#,##0.00' : ci === 4 ? '0.0' : 'General' };
    });
    totRow.height = 24;
  }

  // ── Deal-level detail below totals ────────────────────────────
  const detailStartRow = 6 + months.length + 3;
  ws.mergeCells(`A${detailStartRow}:F${detailStartRow}`);
  const dHdr = ws.getCell(`A${detailStartRow}`);
  dHdr.value = 'DEAL-LEVEL DETAIL — All Active Deals';
  dHdr.style = headerStyle(C.midBlue);
  ws.getRow(detailStartRow).height = 22;

  const dColHdr = ws.getRow(detailStartRow + 1);
  ['Deal Name', 'Stage', 'Amount', 'Weighted Value', 'Probability %', 'Close Date'].forEach((h, i) => {
    const c = dColHdr.getCell(i + 1); c.value = h; c.style = headerStyle(C.darkBlue);
  });
  ws.getRow(detailStartRow + 1).height = 20;

  active.sort((a, b) => (b.amount * b.probability / 100) - (a.amount * a.probability / 100))
    .forEach((d, ri) => {
      const r = ws.getRow(detailStartRow + 2 + ri);
      const prob = d.probability || 0;
      const weighted = (d.amount || 0) * (prob / 100);
      const bg = ri % 2 === 0 ? C.white : C.altRow;
      [d.name, d.stage, d.amount, weighted, prob, d.closeDate].forEach((v, ci) => {
        const c = r.getCell(ci + 1);
        c.value = v;
        c.style = { ...cellStyle(bg, false, ci >= 2 && ci <= 4 ? 'right' : 'left'), numFmt: ci === 2 || ci === 3 ? '#,##0.00' : ci === 4 ? '0.0' : 'General' };
      });
      r.height = 18;
    });
}

// ── ODOO — DEAL VELOCITY ──────────────────────────────────────
// How many days deals spend in each stage — identifies bottlenecks
function buildDealVelocity(wb, pipeline, wonLost) {
  const ws = wb.addWorksheet('⏱️ Deal Velocity', { properties: { tabColor: { argb: 'FF' + C.red } } });
  ws.views = [{ showGridLines: false, state: 'frozen', ySplit: 5 }];
  setColWidths(ws, [28, 14, 18, 18, 20]);

  ws.mergeCells('A2:E2');
  const t = ws.getCell('A2');
  t.value = 'DEAL VELOCITY — SALES CYCLE ANALYSIS';
  t.style = { font: { bold: true, size: 16, name: 'Calibri', color: { argb: 'FF' + C.headerFont } },
              fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.red } },
              alignment: { horizontal: 'center', vertical: 'middle' } };
  ws.getRow(2).height = 40;

  ws.mergeCells('A3:E3');
  const sub = ws.getCell('A3');
  sub.value = 'Based on close dates of Won & Lost deals. Shorter cycle = faster sales. Red = needs attention.';
  sub.style = { font: { size: 10, name: 'Calibri', color: { argb: 'FF' + C.headerFont } },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.midBlue } },
                alignment: { horizontal: 'center', vertical: 'middle' } };
  ws.getRow(3).height = 20;
  ws.getRow(4).height = 8;

  // ── Section 1: Win/Loss by Stage ──────────────────────────────
  const s1 = ws.getRow(5);
  ws.mergeCells('A5:E5');
  s1.getCell(1).value = 'SECTION 1 — Win Rate by Stage';
  s1.getCell(1).style = headerStyle(C.darkBlue);
  ws.getRow(5).height = 22;

  const hRow1 = ws.getRow(6);
  ['Stage', 'Open Deals', 'Won', 'Lost', 'Win Rate %'].forEach((h, i) => {
    const c = hRow1.getCell(i + 1); c.value = h; c.style = headerStyle(C.midBlue);
  });
  ws.getRow(6).height = 20;

  const stageMap = {};
  pipeline.forEach(d => {
    const s = d.stage || 'Unknown';
    if (!stageMap[s]) stageMap[s] = { open: 0, won: 0, lost: 0 };
    stageMap[s].open++;
  });
  (wonLost || []).forEach(d => {
    const s = d.result === 'Won' ? 'Won' : 'Lost';
    if (!stageMap[s]) stageMap[s] = { open: 0, won: 0, lost: 0 };
    if (d.result === 'Won') stageMap[s].won++; else stageMap[s].lost++;
  });

  Object.entries(stageMap).forEach(([stage, agg], ri) => {
    const r = ws.getRow(7 + ri);
    const total = agg.open + agg.won + agg.lost;
    const winRate = total > 0 ? ((agg.won / total) * 100).toFixed(1) : '0.0';
    const bg = ri % 2 === 0 ? C.white : C.altRow;
    [stage, agg.open, agg.won, agg.lost, parseFloat(winRate)].forEach((v, ci) => {
      const c = r.getCell(ci + 1);
      c.value = v;
      c.style = { ...cellStyle(bg, false, ci === 0 ? 'left' : 'right'), numFmt: ci === 4 ? '0.0' : 'General' };
      if (ci === 4 && parseFloat(winRate) >= 50) c.style.font = { ...c.style.font, bold: true, color: { argb: 'FF' + C.green } };
      if (ci === 4 && parseFloat(winRate) < 30 && total > 0) c.style.font = { ...c.style.font, bold: true, color: { argb: 'FF' + C.red } };
    });
    r.height = 18;
  });

  // ── Section 2: Salesperson Velocity ──────────────────────────
  const sec2Row = 7 + Object.keys(stageMap).length + 2;
  ws.mergeCells(`A${sec2Row}:E${sec2Row}`);
  ws.getRow(sec2Row).getCell(1).value = 'SECTION 2 — Salesperson Win Rate';
  ws.getRow(sec2Row).getCell(1).style = headerStyle(C.darkBlue);
  ws.getRow(sec2Row).height = 22;

  const hRow2 = ws.getRow(sec2Row + 1);
  ['Salesperson', 'Open Deals', 'Pipeline Value', 'Won Deals', 'Win Rate %'].forEach((h, i) => {
    const c = hRow2.getCell(i + 1); c.value = h; c.style = headerStyle(C.midBlue);
  });
  ws.getRow(sec2Row + 1).height = 20;

  const spMap2 = {};
  pipeline.forEach(d => {
    const sp = d.pipeline || 'Unassigned';
    if (!spMap2[sp]) spMap2[sp] = { open: 0, pipeVal: 0, won: 0 };
    spMap2[sp].open++;
    spMap2[sp].pipeVal += d.amount || 0;
  });
  (wonLost || []).filter(d => d.result === 'Won').forEach(d => {
    const sp = d.salesperson || 'Unassigned';
    if (!spMap2[sp]) spMap2[sp] = { open: 0, pipeVal: 0, won: 0 };
    spMap2[sp].won++;
  });

  Object.entries(spMap2).sort((a, b) => b[1].pipeVal - a[1].pipeVal).forEach(([sp, agg], ri) => {
    const r = ws.getRow(sec2Row + 2 + ri);
    const total = agg.open + agg.won;
    const winRate = total > 0 ? ((agg.won / total) * 100).toFixed(1) : '0.0';
    const bg = ri % 2 === 0 ? C.white : C.altRow;
    const medal = ri === 0 ? '🥇 ' : ri === 1 ? '🥈 ' : ri === 2 ? '🥉 ' : '';
    [medal + sp, agg.open, agg.pipeVal, agg.won, parseFloat(winRate)].forEach((v, ci) => {
      const c = r.getCell(ci + 1);
      c.value = v;
      c.style = { ...cellStyle(bg, ri < 3, ci === 0 ? 'left' : 'right'), numFmt: ci === 2 ? '#,##0.00' : ci === 4 ? '0.0' : 'General' };
      if (ci === 4 && parseFloat(winRate) >= 50) c.style.font = { ...c.style.font, bold: true, color: { argb: 'FF' + C.green } };
    });
    r.height = 18;
  });
}

// ── ODOO — STAGE ANALYSIS ─────────────────────────────────────
function buildStageAnalysis(wb, pipeline) {
  const ws = wb.addWorksheet('📈 Stage Analysis', { properties: { tabColor: { argb: 'FF' + C.teal } } });
  ws.views = [{ showGridLines: false, state: 'frozen', ySplit: 4 }];
  setColWidths(ws, [28, 14, 20, 16]);

  ws.mergeCells('A2:D2');
  const t = ws.getCell('A2');
  t.value = 'STAGE ANALYSIS';
  t.style = { font: { bold: true, size: 16, name: 'Calibri', color: { argb: 'FF' + C.headerFont } },
              fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.teal } },
              alignment: { horizontal: 'center', vertical: 'middle' } };
  ws.getRow(2).height = 36;
  ws.getRow(3).height = 8;

  const hRow = ws.getRow(4);
  ['Stage', 'Deal Count', 'Pipeline Value', 'Avg Deal Size'].forEach((h, i) => {
    const c = hRow.getCell(i + 1); c.value = h; c.style = headerStyle(C.darkBlue);
  });
  ws.getRow(4).height = 22;

  // Aggregate by stage
  const stageMap = {};
  pipeline.forEach(d => {
    if (!stageMap[d.stage]) stageMap[d.stage] = { count: 0, value: 0 };
    stageMap[d.stage].count++;
    stageMap[d.stage].value += d.amount || 0;
  });
  const stages = Object.entries(stageMap).sort((a, b) => b[1].value - a[1].value);

  stages.forEach(([stage, agg], ri) => {
    const r = ws.getRow(5 + ri);
    const avg = agg.count ? agg.value / agg.count : 0;
    const isWon  = stage.toLowerCase().includes('won');
    const isLost = stage.toLowerCase().includes('lost');
    const bg = ri % 2 === 0 ? C.white : C.altRow;
    [stage, agg.count, agg.value, avg].forEach((v, ci) => {
      const c = r.getCell(ci + 1);
      c.value = v;
      c.style = { ...cellStyle(bg, false, ci === 0 ? 'left' : 'right'), numFmt: ci >= 2 ? '#,##0.00' : 'General' };
      if (isWon  && ci === 0) c.style.font = { ...c.style.font, bold: true, color: { argb: 'FF' + C.green } };
      if (isLost && ci === 0) c.style.font = { ...c.style.font, bold: true, color: { argb: 'FF' + C.red } };
    });
    r.height = 18;
  });

  // Totals row
  if (stages.length) {
    const totRow = ws.getRow(5 + stages.length);
    const totVal = pipeline.reduce((s, d) => s + (d.amount || 0), 0);
    const totAvg = pipeline.length ? totVal / pipeline.length : 0;
    ['TOTAL', pipeline.length, totVal, totAvg].forEach((v, ci) => {
      const c = totRow.getCell(ci + 1);
      c.value = v;
      c.style = { ...headerStyle(C.darkBlue), numFmt: ci >= 2 ? '#,##0.00' : 'General' };
    });
    totRow.height = 22;
  }
}

// ── ODOO — SALESPERSON SHEET ──────────────────────────────────
function buildSalespersonSheet(wb, pipeline, wonLost) {
  const ws = wb.addWorksheet('👤 Salesperson', { properties: { tabColor: { argb: 'FF' + C.purple } } });
  ws.views = [{ showGridLines: false, state: 'frozen', ySplit: 4 }];
  setColWidths(ws, [28, 14, 20, 16, 12]);

  ws.mergeCells('A2:E2');
  const t = ws.getCell('A2');
  t.value = 'SALESPERSON ANALYSIS';
  t.style = { font: { bold: true, size: 16, name: 'Calibri', color: { argb: 'FF' + C.headerFont } },
              fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.purple } },
              alignment: { horizontal: 'center', vertical: 'middle' } };
  ws.getRow(2).height = 36;
  ws.getRow(3).height = 8;

  const hRow = ws.getRow(4);
  ['Salesperson', 'Open Deals', 'Pipeline Value', 'Won Value', 'Win Rate %'].forEach((h, i) => {
    const c = hRow.getCell(i + 1); c.value = h; c.style = headerStyle(C.darkBlue);
  });
  ws.getRow(4).height = 22;

  // Aggregate by salesperson (field: pipeline = user_id name)
  const spMap = {};
  pipeline.forEach(d => {
    const sp = d.pipeline || 'Unassigned';
    if (!spMap[sp]) spMap[sp] = { open: 0, openVal: 0 };
    spMap[sp].open++;
    spMap[sp].openVal += d.amount || 0;
  });
  (wonLost || []).filter(d => d.result === 'Won').forEach(d => {
    const sp = d.salesperson || 'Unassigned';
    if (!spMap[sp]) spMap[sp] = { open: 0, openVal: 0 };
    if (!spMap[sp].wonVal) spMap[sp].wonVal = 0;
    spMap[sp].wonVal += d.amount || 0;
  });

  const rows = Object.entries(spMap).sort((a, b) => (b[1].openVal) - (a[1].openVal));
  rows.forEach(([sp, agg], ri) => {
    const r = ws.getRow(5 + ri);
    const wonVal  = agg.wonVal || 0;
    const winRate = (agg.open + (wonVal > 0 ? 1 : 0)) > 0
      ? ((wonVal > 0 ? 1 : 0) / (agg.open + (wonVal > 0 ? 1 : 0)) * 100).toFixed(1)
      : '0.0';
    const bg = ri % 2 === 0 ? C.white : C.altRow;
    // Gold/silver/bronze top 3
    const medal = ri === 0 ? '🥇 ' : ri === 1 ? '🥈 ' : ri === 2 ? '🥉 ' : '';
    [medal + sp, agg.open, agg.openVal, wonVal, parseFloat(winRate)].forEach((v, ci) => {
      const c = r.getCell(ci + 1);
      c.value = v;
      c.style = { ...cellStyle(bg, ri < 3, ci === 0 ? 'left' : 'right'), numFmt: ci === 2 || ci === 3 ? '#,##0.00' : ci === 4 ? '0.0' : 'General' };
    });
    r.height = 18;
  });

  // Totals
  if (rows.length) {
    const totRow = ws.getRow(5 + rows.length);
    const totOpen = rows.reduce((s, [, a]) => s + a.open, 0);
    const totOpenVal = rows.reduce((s, [, a]) => s + a.openVal, 0);
    const totWon  = rows.reduce((s, [, a]) => s + (a.wonVal || 0), 0);
    ['TOTAL', totOpen, totOpenVal, totWon, ''].forEach((v, ci) => {
      const c = totRow.getCell(ci + 1);
      c.value = v;
      c.style = { ...headerStyle(C.darkBlue), numFmt: ci === 2 || ci === 3 ? '#,##0.00' : 'General' };
    });
    totRow.height = 22;
  }
}

// ── ODOO — CUSTOMER-WISE SHEET ────────────────────────────────
function buildCustomerWiseSheet(wb, pipeline) {
  const ws = wb.addWorksheet('🏢 Customer-wise', { properties: { tabColor: { argb: 'FF' + C.orange } } });
  ws.views = [{ showGridLines: false, state: 'frozen', ySplit: 4 }];
  setColWidths(ws, [32, 14, 20, 14]);

  ws.mergeCells('A2:D2');
  const t = ws.getCell('A2');
  t.value = 'CUSTOMER ANALYSIS';
  t.style = { font: { bold: true, size: 16, name: 'Calibri', color: { argb: 'FF' + C.headerFont } },
              fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.orange } },
              alignment: { horizontal: 'center', vertical: 'middle' } };
  ws.getRow(2).height = 36;
  ws.getRow(3).height = 8;

  const hRow = ws.getRow(4);
  ['Customer / Company', 'Deal Count', 'Pipeline Value', 'Share %'].forEach((h, i) => {
    const c = hRow.getCell(i + 1); c.value = h; c.style = headerStyle(C.darkBlue);
  });
  ws.getRow(4).height = 22;

  const custMap = {};
  pipeline.forEach(d => {
    const cust = d.customer || d.company || 'Unknown';
    if (!custMap[cust]) custMap[cust] = { count: 0, value: 0 };
    custMap[cust].count++;
    custMap[cust].value += d.amount || 0;
  });

  const total = pipeline.reduce((s, d) => s + (d.amount || 0), 0);
  const sorted = Object.entries(custMap).sort((a, b) => b[1].value - a[1].value);

  sorted.forEach(([cust, agg], ri) => {
    const r = ws.getRow(5 + ri);
    const share = total > 0 ? (agg.value / total * 100) : 0;
    const bg = share > 20 ? C.yellow : ri % 2 === 0 ? C.white : C.altRow; // orange flag if >20%
    [cust, agg.count, agg.value, share].forEach((v, ci) => {
      const c = r.getCell(ci + 1);
      c.value = v;
      c.style = { ...cellStyle(bg, share > 20, ci === 0 ? 'left' : 'right'), numFmt: ci === 2 ? '#,##0.00' : ci === 3 ? '0.0' : 'General' };
    });
    r.height = 18;
  });

  if (sorted.length) {
    const totRow = ws.getRow(5 + sorted.length);
    ['TOTAL', pipeline.length, total, 100].forEach((v, ci) => {
      const c = totRow.getCell(ci + 1);
      c.value = v;
      c.style = { ...headerStyle(C.darkBlue), numFmt: ci === 2 ? '#,##0.00' : ci === 3 ? '0.0' : 'General' };
    });
    totRow.height = 22;
  }
}

// ── ODOO — MONTHLY TREND ──────────────────────────────────────
function buildMonthlyTrend(wb, pipeline, wonLost) {
  const ws = wb.addWorksheet('📅 Monthly Trend', { properties: { tabColor: { argb: 'FF' + C.midBlue } } });
  ws.views = [{ showGridLines: false, state: 'frozen', ySplit: 4 }];
  setColWidths(ws, [16, 14, 20, 14, 16]);

  ws.mergeCells('A2:E2');
  const t = ws.getCell('A2');
  t.value = 'MONTHLY TREND';
  t.style = { font: { bold: true, size: 16, name: 'Calibri', color: { argb: 'FF' + C.headerFont } },
              fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.midBlue } },
              alignment: { horizontal: 'center', vertical: 'middle' } };
  ws.getRow(2).height = 36;
  ws.getRow(3).height = 8;

  const hRow = ws.getRow(4);
  ['Month', 'Deals Created', 'Pipeline Value', 'Deals Won', 'Won Value'].forEach((h, i) => {
    const c = hRow.getCell(i + 1); c.value = h; c.style = headerStyle(C.darkBlue);
  });
  ws.getRow(4).height = 22;

  // Group pipeline by close date month
  const monthMap = {};
  const getMonthKey = dateStr => {
    if (!dateStr) return null;
    const d = new Date(dateStr);
    if (isNaN(d)) return null;
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  };

  pipeline.forEach(d => {
    const mk = getMonthKey(d.closeDate);
    if (!mk) return;
    if (!monthMap[mk]) monthMap[mk] = { created: 0, pipeVal: 0, won: 0, wonVal: 0 };
    monthMap[mk].created++;
    monthMap[mk].pipeVal += d.amount || 0;
  });
  (wonLost || []).filter(d => d.result === 'Won').forEach(d => {
    const mk = getMonthKey(d.closeDate);
    if (!mk) return;
    if (!monthMap[mk]) monthMap[mk] = { created: 0, pipeVal: 0, won: 0, wonVal: 0 };
    monthMap[mk].won++;
    monthMap[mk].wonVal += d.amount || 0;
  });

  const months = Object.keys(monthMap).sort();
  months.forEach((mk, ri) => {
    const agg = monthMap[mk];
    const r = ws.getRow(5 + ri);
    const bg = ri % 2 === 0 ? C.white : C.altRow;
    [mk, agg.created, agg.pipeVal, agg.won, agg.wonVal].forEach((v, ci) => {
      const c = r.getCell(ci + 1);
      c.value = v;
      c.style = { ...cellStyle(bg, false, ci === 0 ? 'left' : 'right'), numFmt: ci === 2 || ci === 4 ? '#,##0.00' : 'General' };
    });
    r.height = 18;
  });

  if (months.length) {
    const totRow = ws.getRow(5 + months.length);
    const totCreated = months.reduce((s, m) => s + monthMap[m].created, 0);
    const totPipe    = months.reduce((s, m) => s + monthMap[m].pipeVal, 0);
    const totWon     = months.reduce((s, m) => s + monthMap[m].won, 0);
    const totWonVal  = months.reduce((s, m) => s + monthMap[m].wonVal, 0);
    ['TOTAL', totCreated, totPipe, totWon, totWonVal].forEach((v, ci) => {
      const c = totRow.getCell(ci + 1);
      c.value = v;
      c.style = { ...headerStyle(C.darkBlue), numFmt: ci === 2 || ci === 4 ? '#,##0.00' : 'General' };
    });
    totRow.height = 22;
  }
}

function buildDashboard(wb, data) {
  const ws = wb.addWorksheet('📊 Dashboard', { properties: { tabColor: { argb: 'FF' + C.darkBlue } } });
  ws.views = [{ showGridLines: false }];
  setColWidths(ws, [3, 30, 25, 25, 3]);

  // Title
  ws.mergeCells('B2:D2');
  const title = ws.getCell('B2');
  title.value = `${data.crmName.toUpperCase()} — CRM REPORT`;
  title.style = { font: { bold: true, size: 22, name: 'Calibri', color: { argb: 'FF' + C.headerFont } },
                  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.darkBlue } },
                  alignment: { horizontal: 'center', vertical: 'middle' } };
  ws.getRow(2).height = 48;

  ws.mergeCells('B3:D3');
  const sub = ws.getCell('B3');
  sub.value = `Generated: ${fmtDate()}  |  Developed by Arun Chiramal — Full Stack Developer`;
  sub.style = { font: { bold: false, size: 10, name: 'Calibri', color: { argb: 'FF' + C.headerFont } },
                fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.midBlue } },
                alignment: { horizontal: 'center', vertical: 'middle' } };
  ws.getRow(3).height = 22;
  ws.getRow(4).height = 12;

  // KPI Cards
  const kpis = [
    ['💰 Total Pipeline Value', `$${data.kpis.totalPipeline.toLocaleString('en-US', {minimumFractionDigits:2,maximumFractionDigits:2})}`, C.darkBlue],
    ['📂 Open Deals',           data.kpis.openDeals.toString(),  C.midBlue],
    ['⏰ Closing in 30 Days',   data.kpis.closingSoon.toString(), C.orange],
    ['✅ Won This Month',        data.kpis.wonThisMonth.toString(), '375623'],
    ['📋 Total Contacts',        data.contacts.length.toString(),  C.purple],
    ['🔍 Total Leads',           data.leads.length.toString(),     C.teal],
  ];

  kpis.forEach(([label, value, bg], i) => {
    const row = 5 + i;
    ws.mergeCells(`B${row}:C${row}`);
    const lCell = ws.getCell(`B${row}`);
    lCell.value = label;
    lCell.style = { font: { bold: true, size: 11, name: 'Calibri', color: { argb: 'FF333333' } },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9F0FB' } },
                    alignment: { horizontal: 'left', vertical: 'middle', indent: 1 },
                    border: thinBorder() };

    const vCell = ws.getCell(`D${row}`);
    vCell.value = value;
    vCell.style = { font: { bold: true, size: 14, name: 'Calibri', color: { argb: 'FF' + C.headerFont } },
                    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + bg } },
                    alignment: { horizontal: 'center', vertical: 'middle' },
                    border: thinBorder() };
    ws.getRow(row).height = 30;
  });
}

function buildPipeline(wb, pipeline) {
  const ws = wb.addWorksheet('🔵 Pipeline', { properties: { tabColor: { argb: 'FF' + C.midBlue } } });
  addSheetHeader(ws, '📋 Sales Pipeline', 'All Open Deals', 6);
  setColWidths(ws, [30, 18, 20, 15, 15, 18]);

  const headers = ['Deal Name', 'Amount ($)', 'Stage', 'Close Date', 'Probability %', 'Pipeline/Type'];
  const hRow = ws.getRow(4);
  headers.forEach((h, i) => { const cell = hRow.getCell(i+1); cell.value = h; Object.assign(cell, { style: headerStyle(C.midBlue) }); });
  hRow.height = 20;
  freezeRow(ws, 4);

  pipeline.forEach((d, i) => {
    const row = ws.getRow(5 + i);
    const bg  = i % 2 === 0 ? C.white : C.altRow;
    const stageLower = (d.stage || '').toLowerCase();
    const stageBg = stageLower.includes('won') ? '70AD47' : stageLower.includes('lost') ? 'C00000' : bg;

    row.getCell(1).value = d.name;        row.getCell(1).style = cellStyle(bg);
    row.getCell(2).value = d.amount;      row.getCell(2).style = { ...cellStyle(bg, false, 'right'), numFmt: '#,##0.00' };
    row.getCell(3).value = d.stage;       row.getCell(3).style = { ...cellStyle(stageBg, true, 'center'), font: { bold: true, size: 10, name: 'Calibri', color: { argb: stageLower.includes('won') || stageLower.includes('lost') ? 'FFFFFFFF' : 'FF222222' } } };
    row.getCell(4).value = d.closeDate;   row.getCell(4).style = cellStyle(bg, false, 'center');
    row.getCell(5).value = d.probability ? `${d.probability}%` : ''; row.getCell(5).style = cellStyle(bg, false, 'center');
    row.getCell(6).value = d.pipeline;    row.getCell(6).style = cellStyle(bg);
    row.height = 18;
  });

  if (pipeline.length === 0) { ws.getRow(5).getCell(1).value = 'No pipeline data found.'; }
}

function buildLeads(wb, leads) {
  const ws = wb.addWorksheet('🟡 Leads', { properties: { tabColor: { argb: 'FF' + C.orange } } });
  addSheetHeader(ws, '🔍 Leads', 'All Leads with Status', 6);
  setColWidths(ws, [25, 28, 18, 22, 20, 20]);

  const headers = ['Name', 'Email', 'Phone', 'Company', 'Status', 'Last Activity'];
  const hRow = ws.getRow(4);
  headers.forEach((h, i) => { const cell = hRow.getCell(i+1); cell.value = h; Object.assign(cell, { style: headerStyle(C.orange, C.headerFont) }); });
  hRow.height = 20;
  freezeRow(ws, 4);

  leads.forEach((l, i) => {
    const row = ws.getRow(5 + i);
    const bg  = i % 2 === 0 ? C.white : C.altRow;
    row.getCell(1).value = l.name;         row.getCell(1).style = cellStyle(bg, true);
    row.getCell(2).value = l.email;        row.getCell(2).style = cellStyle(bg);
    row.getCell(3).value = l.phone;        row.getCell(3).style = cellStyle(bg);
    row.getCell(4).value = l.company;      row.getCell(4).style = cellStyle(bg);
    row.getCell(5).value = l.status;       row.getCell(5).style = cellStyle(bg, false, 'center');
    row.getCell(6).value = l.lastActivity ? String(l.lastActivity).substring(0,10) : ''; row.getCell(6).style = cellStyle(bg, false, 'center');
    row.height = 18;
  });

  if (leads.length === 0) { ws.getRow(5).getCell(1).value = 'No leads data found.'; }
}

function buildContacts(wb, contacts) {
  const ws = wb.addWorksheet('🟢 Contacts', { properties: { tabColor: { argb: 'FF' + C.green } } });
  addSheetHeader(ws, '👤 Contacts', 'Full Contact List with Last Activity', 6);
  setColWidths(ws, [25, 28, 18, 22, 20, 20]);

  const headers = ['Name', 'Email', 'Phone', 'Company', 'Title/Role', 'Last Activity'];
  const hRow = ws.getRow(4);
  headers.forEach((h, i) => { const cell = hRow.getCell(i+1); cell.value = h; Object.assign(cell, { style: headerStyle(C.green, C.headerFont) }); });
  hRow.height = 20;
  freezeRow(ws, 4);

  contacts.forEach((c, i) => {
    const row = ws.getRow(5 + i);
    const bg  = i % 2 === 0 ? C.white : C.altRow;
    row.getCell(1).value = c.name;         row.getCell(1).style = cellStyle(bg, true);
    row.getCell(2).value = c.email;        row.getCell(2).style = cellStyle(bg);
    row.getCell(3).value = c.phone;        row.getCell(3).style = cellStyle(bg);
    row.getCell(4).value = c.company;      row.getCell(4).style = cellStyle(bg);
    row.getCell(5).value = c.title;        row.getCell(5).style = cellStyle(bg);
    row.getCell(6).value = c.lastActivity ? String(c.lastActivity).substring(0,10) : ''; row.getCell(6).style = cellStyle(bg, false, 'center');
    row.height = 18;
  });

  if (contacts.length === 0) { ws.getRow(5).getCell(1).value = 'No contacts data found.'; }
}

function buildActivities(wb, activities) {
  const ws = wb.addWorksheet('🔔 Activities', { properties: { tabColor: { argb: 'FF' + C.purple } } });
  addSheetHeader(ws, '📅 Activities & Tasks', 'Due Today, This Week, Overdue', 4);
  setColWidths(ws, [35, 20, 20, 20]);

  const headers = ['Subject', 'Status', 'Due Date', 'Type'];
  const hRow = ws.getRow(4);
  headers.forEach((h, i) => { const cell = hRow.getCell(i+1); cell.value = h; Object.assign(cell, { style: headerStyle(C.purple) }); });
  hRow.height = 20;
  freezeRow(ws, 4);

  activities.forEach((a, i) => {
    const row = ws.getRow(5 + i);
    const bg  = i % 2 === 0 ? C.white : C.altRow;
    let dueBg = bg;
    if (a.dueDate) {
      const days = (new Date(a.dueDate) - new Date()) / 86400000;
      if (days < 0) dueBg = 'FFCCCC'; // overdue — red
      else if (days <= 7) dueBg = C.yellow; // due this week — yellow
    }
    row.getCell(1).value = a.subject;  row.getCell(1).style = cellStyle(bg);
    row.getCell(2).value = a.status;   row.getCell(2).style = cellStyle(bg, false, 'center');
    row.getCell(3).value = a.dueDate ? String(a.dueDate).substring(0,10) : ''; row.getCell(3).style = cellStyle(dueBg, false, 'center');
    row.getCell(4).value = a.type;     row.getCell(4).style = cellStyle(bg);
    row.height = 18;
  });

  if (activities.length === 0) { ws.getRow(5).getCell(1).value = 'No activities data found.'; }
}

function buildWonLost(wb, wonLost) {
  const ws = wb.addWorksheet('🏆 Won & Lost', { properties: { tabColor: { argb: 'FF375623' } } });
  addSheetHeader(ws, '🏆 Won & Lost Deals', 'Closed Deals This Month', 4);
  setColWidths(ws, [30, 20, 15, 20]);

  const headers = ['Deal Name', 'Amount ($)', 'Result', 'Close Date'];
  const hRow = ws.getRow(4);
  headers.forEach((h, i) => { const cell = hRow.getCell(i+1); cell.value = h; Object.assign(cell, { style: headerStyle('375623') }); });
  hRow.height = 20;
  freezeRow(ws, 4);

  wonLost.forEach((d, i) => {
    const row = ws.getRow(5 + i);
    const bg  = i % 2 === 0 ? C.white : C.altRow;
    const resultBg = d.result === 'Won' ? 'D9EAD3' : 'FFCCCC';
    row.getCell(1).value = d.name;     row.getCell(1).style = cellStyle(bg, true);
    row.getCell(2).value = d.amount;   row.getCell(2).style = { ...cellStyle(bg, false, 'right'), numFmt: '#,##0.00' };
    row.getCell(3).value = d.result;   row.getCell(3).style = cellStyle(resultBg, true, 'center');
    row.getCell(4).value = d.closeDate ? String(d.closeDate).substring(0,10) : ''; row.getCell(4).style = cellStyle(bg, false, 'center');
    row.height = 18;
  });

  if (wonLost.length === 0) { ws.getRow(5).getCell(1).value = 'No won/lost deals this period.'; }
}

// ── MAIN HANDLER ──────────────────────────────────────────────
const SUPPORTED_CRMS = ['erpnext', 'odoo', 'salesforce', 'pipedrive', 'freshsales', 'zoho'];

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  // Only allow POST
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed. Use POST.' });

  try {
    const { crm, apiKey, domain, username, password, instanceUrl, accessToken, database, reportType, licenseKey, fiscalYear } = req.body || {};

    if (!crm)        return res.status(400).json({ error: 'CRM type required.' });
    if (!reportType) return res.status(400).json({ error: 'Report type required.' });

    // Validate CRM is supported
    if (!SUPPORTED_CRMS.includes(crm.toLowerCase())) {
      return res.status(400).json({ error: `Unsupported CRM: ${crm}. Supported: ${SUPPORTED_CRMS.join(', ')}` });
    }

    // Validate license key
    if (!licenseKey) {
      reportError({ product: 'CRM Data Extractor', version: 'v2.0', message: 'License key missing (401)', crm: crm || '', extra: 'stage:license-check' }).catch(() => {});
      return res.status(401).json({ error: 'License key required.' });
    }

    // Validate required credentials per CRM before calling fetch functions
    const c = crm.toLowerCase();
    if ((c === 'erpnext' || c === 'odoo' || c === 'freshsales') && !domain)
      return res.status(400).json({ error: 'URL / domain is required.' });
    if ((c === 'erpnext' || c === 'odoo') && (!username || !password))
      return res.status(400).json({ error: 'Username and password are required.' });
    if ((c === 'hubspot' || c === 'zoho' || c === 'pipedrive') && !apiKey)
      return res.status(400).json({ error: 'API key is required.' });
    if (c === 'salesforce' && (!instanceUrl || !accessToken))
      return res.status(400).json({ error: 'Instance URL and access token are required.' });

    let data;
    switch (c) {
      case 'hubspot':     data = await fetchHubSpot(apiKey, reportType); break;
      case 'zoho':        data = await fetchZoho(apiKey, reportType); break;
      case 'pipedrive':   data = await fetchPipedrive(apiKey, reportType); break;
      case 'freshsales':  data = await fetchFreshsales(apiKey, domain, reportType); break;
      case 'erpnext':     data = await fetchERPNext(domain, username, password, reportType, fiscalYear); break;
      case 'salesforce':  data = await fetchSalesforce(instanceUrl, accessToken, reportType); break;
      case 'odoo':        data = await fetchOdoo(domain, database, username, password, reportType, fiscalYear); break;
      default:            return res.status(400).json({ error: `Unknown CRM: ${crm}` });
    }

    const buffer = await buildExcel(data);
    const dateStr = new Date().toISOString().split('T')[0];
    const rtSuffix = reportType && reportType !== 'full' ? `_${reportType}` : '';
    const filename = `CRM_Report_${data.crmName.replace(/\s/g,'_')}${rtSuffix}_${dateStr}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', buffer.length);
    res.setHeader('Access-Control-Expose-Headers', 'X-CRM-Warnings');
    if (data.warnings && data.warnings.length > 0) {
      // Strip non-ASCII to keep header value safe across all CRMs
      const safeWarning = `${data.warnings.length} endpoint(s) had issues - check Warnings sheet`
        .replace(/[^\x20-\x7E]/g, '');
      res.setHeader('X-CRM-Warnings', safeWarning);
    }
    return res.status(200).send(Buffer.from(buffer));

  } catch (err) {
    console.error('CRM fetch error:', err.message);
    const msg = err.message || 'Failed to fetch CRM data.';
    const lower = msg.toLowerCase();

    // Silently report to Google Sheet — never await, never block
    const { crm, reportType, licenseKey } = req.body || {};
    reportError({
      product: 'CRM Data Extractor',
      version: 'v2.0',
      message: msg,
      key:     licenseKey,
      crm:     crm ? `${crm}/${reportType || ''}` : '',
      extra:   `method:${req.method}`,
    }).catch(() => {});

    if (lower.includes('login failed') || lower.includes('check your username') || lower.includes('authentication') || lower.includes('401') || lower.includes('403'))
      return res.status(401).json({ error: msg });
    if (lower.includes('cannot reach') || lower.includes('econnrefused') || lower.includes('enotfound'))
      return res.status(503).json({ error: msg });
    return res.status(500).json({ error: msg });
  }
};
