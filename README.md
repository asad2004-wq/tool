<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title> Invoice Dashboard</title>

<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

<style>
body {
  background: linear-gradient(135deg,#eef2f7,#dfe9f3);
  font-family: 'Segoe UI';
}
.container { margin-top:40px; max-width:1350px; }

.card {
  border-radius:20px;
  padding:25px;
  box-shadow:0 10px 30px rgba(0,0,0,0.1);
}

.heading {
  font-size:28px;
  font-weight:bold;
  text-align:center;
  margin-bottom:20px;
  color:#0d6efd;
}

.stat-card {
  border-radius:12px;
  padding:15px;
  font-weight:600;
  font-size:14px;
  color:white;
  transition:0.3s;
}
.stat-card:hover {
  transform:translateY(-3px);
}

.sale { background:linear-gradient(45deg,#28a745,#5ddf8c); }
.credit { background:linear-gradient(45deg,#17a2b8,#63d9e8); }
.net { background:linear-gradient(45deg,#343a40,#6c757d); }
</style>
</head>

<body>

<div class="container">
<div class="card">

<div class="heading">Invoice Dashboard</div>

<label class="fw-bold">Upload FBR Files</label>
<input type="file" id="fbrFiles" class="form-control mb-3" multiple>

<label class="fw-bold">Upload POS Files</label>
<input type="file" id="posFiles" class="form-control mb-3" multiple>

<button class="btn btn-primary w-100 mb-2" onclick="processFiles()">Process</button>
<button class="btn btn-success w-100 mb-3" onclick="downloadExcel()">Download Excel</button>

<div id="dashboard" style="display:none;">

<h5>📊 Overview</h5>
<div class="row mb-3 text-center" id="topStats"></div>

<h5>💰 Balance Summary</h5>
<div class="row mb-3 text-center" id="balanceStats"></div>

<h5>🧾 Tax Summary</h5>
<div class="row mb-4 text-center" id="taxStats"></div>

<div class="row mb-3">
<div class="col-md-6">
<input type="text" id="searchInput" class="form-control" placeholder="Search Invoice / POSID">
</div>
<div class="col-md-3">
<select id="filterType" class="form-control">
<option value="">All</option>
<option value="FBR">Missing in FBR</option>
<option value="POS">Missing in POS</option>
</select>
</div>
</div>

<div id="tableContainer"></div>

</div>
</div>
</div>

<script>

let finalMismatches = [];

function convertExcelDate(serial) {
  if (!serial || isNaN(serial)) return serial;

  let utc_days = Math.floor(serial - 25569);
  let utc_value = utc_days * 86400;
  let date_info = new Date(utc_value * 1000);

  let fractional_day = serial - Math.floor(serial);
  let total_seconds = Math.round(fractional_day * 86400);

  let hours = Math.floor(total_seconds / 3600);
  let minutes = Math.floor((total_seconds % 3600) / 60);

  return `${date_info.getFullYear()}-${String(date_info.getMonth()+1).padStart(2,'0')}-${String(date_info.getDate()).padStart(2,'0')} ${String(hours).padStart(2,'0')}:${String(minutes).padStart(2,'0')}`;
}

function readExcel(file) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      resolve(XLSX.utils.sheet_to_json(sheet));
    };
    reader.readAsArrayBuffer(file);
  });
}

async function processFiles() {

  finalMismatches = [];

  const fbrFiles = Array.from(document.getElementById('fbrFiles').files);
  const posFiles = Array.from(document.getElementById('posFiles').files);

  let fbrData = [], posData = [];

  for (const f of fbrFiles) fbrData.push(...await readExcel(f));
  for (const f of posFiles) posData.push(...await readExcel(f));

  const fbrMap = new Map();
  const posMap = new Map();

  fbrData.forEach(r=>{
    let inv = (r['FBRInvoiceNumber']||'').toString().trim();
    if(inv) fbrMap.set(inv,{
      InvoiceNo:inv,
      Type:r['InvoiceType'],
      DateTime:convertExcelDate(r['DateTime']),
      Balance:r['Total_Balance'],
      TaxCharged:r['Tax_Charged'],
      Discount:Number(r['Discount'])||0,
      POSID:r['POS_ID']
    });
  });

  posData.forEach(r=>{
    let inv = (r['FBR Number']||'').toString().trim();
    if(inv) posMap.set(inv,{
      InvoiceNo:inv,
      Type:r['Transaction Type'],
      DateTime:isNaN(r['Voucher Date']) ? r['Voucher Date'] : convertExcelDate(r['Voucher Date']),
      Balance:r['Sale Value'],
      TaxCharged:r['TaxCharged'],
      Discount:Number(r['Total Discount'])||0,
      POSID:r['FBR POS ID']
    });
  });

  fbrMap.forEach((r,inv)=>{ if(!posMap.has(inv)) finalMismatches.push({...r,MissingIn:'POS'}); });
  posMap.forEach((r,inv)=>{ if(!fbrMap.has(inv)) finalMismatches.push({...r,MissingIn:'FBR'}); });

  renderDashboard();
}

function renderDashboard(){
  document.getElementById("dashboard").style.display="block";
  renderStats();
  renderTable();
}

function renderStats(){

  let total = finalMismatches.length;
  let fbr = finalMismatches.filter(x => x.MissingIn === 'FBR').length;
  let pos = finalMismatches.filter(x => x.MissingIn === 'POS').length;

  let saleBal=0, creditBal=0, saleTax=0, creditTax=0;

  finalMismatches.forEach(x=>{
    let t=(x.Type||'').toUpperCase();
    let b=Number(x.Balance)||0;
    let tax=Number(x.TaxCharged)||0;

    if(t.includes('SALE')){ saleBal+=b; saleTax+=tax; }
    else if(t.includes('CREDIT')){ creditBal+=b; creditTax+=tax; }
  });

  let netBal = saleBal - creditBal;
  let netTax = saleTax - creditTax;

  document.getElementById("topStats").innerHTML=`
    <div class="col-md-4"><div class="stat-card bg-primary">Total<br>${total}</div></div>
    <div class="col-md-4"><div class="stat-card bg-danger">FBR Missing<br>${fbr}</div></div>
    <div class="col-md-4"><div class="stat-card bg-warning text-dark">POS Missing<br>${pos}</div></div>
  `;

  document.getElementById("balanceStats").innerHTML=`
    <div class="col-md-4"><div class="stat-card sale">SALE<br>${saleBal.toFixed(2)}</div></div>
    <div class="col-md-4"><div class="stat-card credit">CREDIT<br>${creditBal.toFixed(2)}</div></div>
    <div class="col-md-4"><div class="stat-card net">NET<br>${netBal.toFixed(2)}</div></div>
  `;

  document.getElementById("taxStats").innerHTML=`
    <div class="col-md-4"><div class="stat-card sale">SALE TAX<br>${saleTax.toFixed(2)}</div></div>
    <div class="col-md-4"><div class="stat-card credit">CREDIT TAX<br>${creditTax.toFixed(2)}</div></div>
    <div class="col-md-4"><div class="stat-card net">NET TAX<br>${netTax.toFixed(2)}</div></div>
  `;
}

function renderTable(){
  let search = (document.getElementById("searchInput").value || '').toLowerCase();
  let filter = document.getElementById("filterType").value;

  let data = finalMismatches.filter(r => {
    let invoice = (r.InvoiceNo || '').toString().toLowerCase();
    let posid = (r.POSID || '').toString().toLowerCase();

    return (invoice.includes(search) || posid.includes(search)) &&
           (!filter || r.MissingIn === filter);
  });

  let html=`<table class="table table-bordered table-hover">
  <thead class="table-dark">
  <tr>
  <th>Invoice</th><th>Missing</th><th>Type</th><th>Date</th><th>Balance</th><th>Tax</th><th>Discount</th><th>POSID</th>
  </tr></thead><tbody>`;

  data.forEach(r=>{
    html+=`<tr>
    <td>${r.InvoiceNo}</td>
    <td>${r.MissingIn}</td>
    <td>${r.Type}</td>
    <td>${r.DateTime}</td>
    <td>${r.Balance}</td>
    <td>${r.TaxCharged}</td>
    <td>${r.Discount}</td>
    <td>${r.POSID}</td>
    </tr>`;
  });

  html+="</tbody></table>";
  document.getElementById("tableContainer").innerHTML=html;
}

document.addEventListener("input", e=>{
  if(e.target.id==="searchInput"||e.target.id==="filterType") renderTable();
});

function downloadExcel(){

  const wb = XLSX.utils.book_new();

  let ws = XLSX.utils.json_to_sheet(finalMismatches);
  ws['!cols'] = Object.keys(finalMismatches[0]||{}).map(()=>({wch:20}));
  XLSX.utils.book_append_sheet(wb, ws, "All Data");

  let summary={};
  finalMismatches.forEach(r=>{
    if(!summary[r.POSID]) summary[r.POSID]={count:0,balance:0};
    summary[r.POSID].count++;
    summary[r.POSID].balance+=Number(r.Balance)||0;
  });

  let arr=[];
  for(let p in summary){
    arr.push({POSID:p,Total:summary[p].count,Balance:summary[p].balance});
  }

  let ws2=XLSX.utils.json_to_sheet(arr);
  XLSX.utils.book_append_sheet(wb, ws2, "POS Summary");

  let grouped={};
  finalMismatches.forEach(r=>{
    if(!grouped[r.POSID]) grouped[r.POSID]=[];
    grouped[r.POSID].push(r);
  });

  for(let p in grouped){
    let ws=XLSX.utils.json_to_sheet(grouped[p]);
    XLSX.utils.book_append_sheet(wb, ws, p.substring(0,30));
  }

  XLSX.writeFile(wb,"Premium_Report.xlsx");
}

</script>

</body>
</html>
