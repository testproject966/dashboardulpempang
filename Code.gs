/*******************************************************
 * EXECUTIVE DASHBOARD
 * PLN ULP EMPANG | PLN UP3 SUMBAWA
 * Google Apps Script + Google Spreadsheet
 *
 * VERSI PRESENTASI MANAJEMEN
 * - Executive Dashboard
 * - Responsive Desktop & Android
 * - Filter periode & bidang
 * - KPI per bidang
 * - Trend bulanan
 * - Status pekerjaan
 * - Target vs realisasi pengusahaan
 * - Pengumuman / highlight
 * - Detail data
 * - Input & edit data
 *******************************************************/

const DB_SPREADSHEET_ID = '1UmBR6Um7WuF8CwsKCxnwVa1tKjyGTyzKxurUBy_b5fo';

function db_(){
  return SpreadsheetApp.openById(DB_SPREADSHEET_ID);
}

const CFG = {
  UNIT: 'PLN ULP EMPANG',
  PARENT: 'PLN UP3 SUMBAWA',
  TZ: 'Asia/Makassar',
  SHEETS: {
    USERS: 'Users',
    TEKNIK: 'Teknik Jaringan',
    TRANSAKSI: 'Transaksi Energi',
    PELAYANAN: 'Pelayanan Pelanggan',
    K3L: 'K3L',
    PENGUSAHAAN: 'Pengusahaan',
    PENGUMUMAN: 'Pengumuman'
  }
};

const SCHEMA = {
  'Users':['Email','Nama','Bidang','Role','Aktif'],
  'Teknik Jaringan':['ID','Tanggal','Penyulang','Lokasi','Jenis Pekerjaan','Uraian','Status','Petugas','Prioritas','Foto','Keterangan'],
  'Transaksi Energi':['ID','Tanggal','ID Pelanggan','Nomor Meter','Merk Meter','Jenis Meter','Kegiatan','Lokasi','Petugas','Status','Temuan','Foto','Keterangan'],
  'Pelayanan Pelanggan':['ID','Tanggal','ID Pelanggan','Jenis Pelayanan','Nama Pelanggan','Lokasi','Status','Petugas','SLA','Keterangan'],
  'K3L':['ID','Tanggal','Kegiatan','Lokasi','Petugas','Status','Temuan','Tindak Lanjut','Foto','Keterangan'],
  'Pengusahaan':['ID','Tanggal','Indikator','Target','Realisasi','Satuan','Persentase','Keterangan'],
  'Pengumuman':['ID','Tanggal','Judul','Isi','Bidang','Aktif']
};

function doGet(e){
  const p=(e&&e.parameter)||{};
  try{
    // API untuk GitHub Pages / Android. Backend dan Spreadsheet tetap di Apps Script.
    if(String(p.api||'')==='health'){
      const ss=db_();
      const sheets=ss.getSheets().map(sh=>({name:sh.getName(),rows:Math.max(0,sh.getLastRow()-1),columns:sh.getLastColumn()}));
      return jsonp_({ok:true,service:'BRIGHT ULP EMPANG API',spreadsheet:ss.getName(),checkedAt:Utilities.formatDate(new Date(),CFG.TZ,'yyyy-MM-dd HH:mm:ss'),sheets:sheets},p.callback);
    }
    if(String(p.api||'')==='getData'){
      const d=getExecutive({year:p.year||'',month:p.month||'',status:p.status||''});
      return jsonp_(d,p.callback);
    }
    if(String(p.api||'')==='getModule'){
      const filters={year:p.year||'',month:p.month||'',status:p.status||''};
      const rows=getModule(String(p.bidang||''),filters);
      return jsonp_(rows,p.callback);
    }
    if(String(p.api||'')==='getPengusahaanDashboard'){
      const d=getPengusahaanDashboard({year:p.year||'',month:p.month||'',status:p.status||''});
      return jsonp_(d,p.callback);
    }
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Executive Dashboard | PLN ULP Empang')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }catch(err){ return jsonp_({ok:false,error:String(err)},p.callback); }
}
function doPost(e){
  try{
    const p=JSON.parse((e&&e.postData&&e.postData.contents)||'{}');
    if(p.api==='appendRows') return jsonp_({ok:true,result:appendRows(String(p.sheet||''),p.rows||[])},p.callback);
    if(p.api==='addRecord') return jsonp_({ok:true,result:addRecord(String(p.sheet||''),p.row||[])},p.callback);
    return jsonp_({ok:false,error:'API POST tidak dikenal.'},p.callback);
  }catch(err){ return jsonp_({ok:false,error:String(err)},(e&&e.parameter&&e.parameter.callback)||''); }
}

function apiSheetName_(name){
  const map={
    teknik:'Teknik Jaringan',jaringan:'Teknik Jaringan',
    transaksi:'Transaksi Energi',transaksi_energi:'Transaksi Energi',
    pelayanan:'Pelayanan Pelanggan',pelayanan_pelanggan:'Pelayanan Pelanggan',
    k3l:'K3L',pengusahaan:'Pengusahaan',
    KPI:'KPI',kpi:'KPI',
    PENJUALAN:'PENJUALAN','PLGTARIF':'PLG/TARIF','PLGKATEGORI':'PLG/KATEGORI',
    'PLGKEC':'PLG/KEC','PLGKECAMATAN':'PLG/KEC','ASET':'ASET','SDM':'SDM','TAD':'TAD'
  };
  return map[name]||map[String(name).toLowerCase()]||name;
}

function appendRows(sheetName,rows){
  const name=apiSheetName_(sheetName);
  if(!Array.isArray(rows)||!rows.length) return {count:0,sheet:name};
  const ss=db_();
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  const existingHeaders=sh.getLastColumn()>0&&sh.getLastRow()>0?sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]:[];
  if(!existingHeaders.length || existingHeaders.every(x=>x==='')){
    sh.getRange(1,1,1,rows[0].length).setValues([rows[0].map((x,i)=>String(x||('Kolom '+(i+1))) )]);
    rows=rows.slice(1);
  }
  if(rows.length){
    const width=sh.getLastColumn();
    const normalized=rows.map(r=>Array.from({length:width},(_,i)=>r[i]??''));
    sh.getRange(sh.getLastRow()+1,1,normalized.length,width).setValues(normalized);
  }
  return {count:rows.length,sheet:name};
}

function addRecord(sheetName,row){
  const name=apiSheetName_(sheetName);
  const ss=db_();
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  if(sh.getLastColumn()===0){
    sh.getRange(1,1,1,row.length).setValues([row.map((_,i)=>'Kolom '+(i+1))]);
  }
  const width=sh.getLastColumn();
  sh.getRange(sh.getLastRow()+1,1,1,width).setValues([Array.from({length:width},(_,i)=>row[i]??'')]);
  return {ok:true,sheet:name};
}

function jsonp_(data,callback){
  const cb=String(callback||'').replace(/[^A-Za-z0-9_.$]/g,'');
  const body=JSON.stringify(data);
  if(cb) return ContentService.createTextOutput(cb+'('+body+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function setupDatabase(){
  const ss=db_();
  Object.keys(SCHEMA).forEach(name=>{
    let sh=ss.getSheetByName(name);
    if(!sh) sh=ss.insertSheet(name);
    if(sh.getLastRow()===0) sh.getRange(1,1,1,SCHEMA[name].length).setValues([SCHEMA[name]]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,SCHEMA[name].length)
      .setFontWeight('bold').setBackground('#102A43').setFontColor('#FFFFFF');
  });

  const users=ss.getSheetByName(CFG.SHEETS.USERS);
  if(users.getLastRow()===1){
    users.getRange(2,1,3,5).setValues([
      ['','Administrator ULP','Semua','ADMIN','YA'],
      ['','Teknik Jaringan','Teknik Jaringan','PETUGAS','YA'],
      ['','K3L','K3L','PETUGAS','YA']
    ]);
  }

  const p=ss.getSheetByName(CFG.SHEETS.PENGUMUMAN);
  if(p.getLastRow()===1){
    p.getRange(2,1,3,6).setValues([
      ['PGM001',new Date(),'Prioritas Keselamatan','Pastikan seluruh pekerjaan dilaksanakan sesuai standar keselamatan kerja.','K3L','YA'],
      ['PGM002',new Date(),'Monitoring Kinerja','Perbarui realisasi pekerjaan dan indikator unit secara berkala.','Semua','YA'],
      ['PGM003',new Date(),'Executive Dashboard','Dashboard merupakan ringkasan monitoring unit kerja.','Semua','YA']
    ]);
  }

  seedDemo_();
  return 'Database berhasil disiapkan.';
}

function getUser_(){
  const email=Session.getActiveUser().getEmail()||'';
  const rows=read_(CFG.SHEETS.USERS);
  const u=rows.find(r=>email && String(r.Email||'').toLowerCase()===email.toLowerCase() && String(r.Aktif||'').toUpperCase()==='YA');
  return u?{email,nama:u.Nama,bidang:u.Bidang,role:u.Role}:{email,nama:email?email.split('@')[0]:'Pengguna',bidang:'Semua',role:'VIEWER'};
}

function getExecutive(filters){
  filters=filters||{};
  const rows={
    teknik:filter_(read_(CFG.SHEETS.TEKNIK),filters),
    transaksi:filter_(read_(CFG.SHEETS.TRANSAKSI),filters),
    pelayanan:filter_(read_(CFG.SHEETS.PELAYANAN),filters),
    k3l:filter_(read_(CFG.SHEETS.K3L),filters),
    pengusahaan:filter_(read_(CFG.SHEETS.PENGUSAHAAN),filters)
  };

  const summary={
    teknik:summary_(rows.teknik),
    transaksi:summary_(rows.transaksi),
    pelayanan:summary_(rows.pelayanan),
    k3l:summary_(rows.k3l),
    pengusahaan:business_(rows.pengusahaan)
  };

  return {
    unit:CFG.UNIT,parent:CFG.PARENT,
    updated:Utilities.formatDate(new Date(),CFG.TZ,'dd MMMM yyyy HH:mm'),
    user:getUser_(),rows,summary,
    trend:trend_(filters),
    latest:latest_(rows),
    announcements:read_(CFG.SHEETS.PENGUMUMAN).filter(r=>String(r.Aktif||'').toUpperCase()==='YA').slice(-6).reverse(),
    years:years_(),months:months_()
  };
}

function filter_(rows,f){
  return rows.filter(r=>{
    const d=String(r.Tanggal||'');
    if(f.year && d.slice(0,4)!==String(f.year)) return false;
    if(f.month && Number(d.slice(5,7))!==Number(f.month)) return false;
    if(f.status && f.status!=='Semua' && String(r.Status||'')!==String(f.status)) return false;
    return true;
  });
}

function summary_(rows){
  let s={total:rows.length,selesai:0,proses:0,belum:0,lainnya:0};
  rows.forEach(r=>{
    const x=String(r.Status||'').toLowerCase();
    if(x.includes('selesai'))s.selesai++;
    else if(x.includes('proses'))s.proses++;
    else if(x.includes('belum'))s.belum++;
    else s.lainnya++;
  });
  s.rate=s.total?s.selesai/s.total*100:0;
  return s;
}

function business_(rows){
  let t=0,r=0;
  rows.forEach(x=>{t+=Number(x.Target)||0;r+=Number(x.Realisasi)||0});
  return {total:rows.length,target:t,realisasi:r,rate:t?r/t*100:0};
}

function trend_(f){
  const all={
    teknik:read_(CFG.SHEETS.TEKNIK),
    transaksi:read_(CFG.SHEETS.TRANSAKSI),
    pelayanan:read_(CFG.SHEETS.PELAYANAN),
    k3l:read_(CFG.SHEETS.K3L)
  };
  const year=Number(f.year||new Date().getFullYear());
  const labels=['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  return labels.map((label,i)=>{
    let total=0,done=0,process=0,pending=0;
    Object.keys(all).forEach(k=>{
      all[k].forEach(r=>{
        const d=String(r.Tanggal||'');
        if(Number(d.slice(0,4))===year && Number(d.slice(5,7))===i+1){
          total++;
          const s=String(r.Status||'').toLowerCase();
          if(s.includes('selesai'))done++;
          else if(s.includes('proses'))process++;
          else if(s.includes('belum'))pending++;
        }
      });
    });
    return {label,total,done,process,pending};
  });
}

function latest_(all){
  const names={teknik:'Teknik Jaringan',transaksi:'Transaksi Energi',pelayanan:'Pelayanan Pelanggan',k3l:'K3L',pengusahaan:'Pengusahaan'};
  let a=[];
  Object.keys(all).forEach(k=>all[k].forEach(r=>a.push({
    tanggal:r.Tanggal||'',bidang:names[k],
    kegiatan:r['Jenis Pekerjaan']||r.Kegiatan||r['Jenis Pelayanan']||r.Indikator||'-',
    lokasi:r.Lokasi||'-',status:r.Status||'-'
  })));
  return a.sort((a,b)=>String(b.tanggal).localeCompare(String(a.tanggal))).slice(0,12);
}

function years_(){
  const set={};
  set[String(new Date().getFullYear())]=true;
  Object.keys(SCHEMA).forEach(n=>{
    if(n==='Users'||n==='Pengumuman')return;
    read_(n).forEach(r=>{const y=String(r.Tanggal||'').slice(0,4);if(y)set[y]=true;});
  });
  return Object.keys(set).sort().reverse();
}

function months_(){
  return ['Semua Bulan','Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember']
    .map((t,i)=>({v:i,t}));
}

function getModule(bidang,filters){
  const map={teknik:CFG.SHEETS.TEKNIK,transaksi:CFG.SHEETS.TRANSAKSI,pelayanan:CFG.SHEETS.PELAYANAN,k3l:CFG.SHEETS.K3L,pengusahaan:CFG.SHEETS.PENGUSAHAAN};
  if(!map[bidang])throw new Error('Bidang tidak ditemukan.');
  return filter_(read_(map[bidang]),filters||{});
}

function saveRecord(p){
  const map={teknik:CFG.SHEETS.TEKNIK,transaksi:CFG.SHEETS.TRANSAKSI,pelayanan:CFG.SHEETS.PELAYANAN,k3l:CFG.SHEETS.K3L,pengusahaan:CFG.SHEETS.PENGUSAHAAN};
  const sh=db_().getSheetByName(map[p.bidang]);
  if(!sh)throw new Error('Sheet tidak ditemukan.');
  const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const row=h.map(x=>p.values[x]??'');
  const id=h.indexOf('ID'),date=h.indexOf('Tanggal');
  if(id>=0&&!row[id])row[id]=p.bidang.toUpperCase()+'-'+Utilities.getUuid().slice(0,8);
  if(date>=0&&row[date])row[date]=new Date(row[date]);
  if(p.mode==='edit'){
    const data=sh.getDataRange().getValues();
    for(let i=1;i<data.length;i++)if(String(data[i][id])===String(p.id)){sh.getRange(i+1,1,1,row.length).setValues([row]);return {ok:true,message:'Data diperbarui.'};}
    throw new Error('Data tidak ditemukan.');
  }
  sh.appendRow(row);
  return {ok:true,message:'Data ditambahkan.'};
}


/* =====================================================
 * MODUL DATA PENGUSAHAAN
 * Membaca sheet existing:
 * PENJUALAN, PLG/TARIF, PLG/KATEGORI, PLG/KEC,
 * ASET, SDM, TAD
 * Modul ini terpisah dari Executive Dashboard utama.
 * ===================================================== */

const PENGUSAHAAN_DATA_SHEETS = {
  penjualan:'PENJUALAN',
  tarif:'PLG/TARIF',
  kategori:'PLG/KATEGORI',
  kec:'PLG/KEC',
  aset:'ASET',
  sdm:'SDM',
  tad:'TAD'
};

function getPengusahaanDashboard(filters){
  filters=filters||{};
  const penjualan = filterPengusahaanPeriod_(read_(PENGUSAHAAN_DATA_SHEETS.penjualan), filters);
  const tarif = filterPengusahaanPeriod_(read_(PENGUSAHAAN_DATA_SHEETS.tarif), filters);
  const kategori = filterPengusahaanPeriod_(read_(PENGUSAHAAN_DATA_SHEETS.kategori), filters);
  const kec = read_(PENGUSAHAAN_DATA_SHEETS.kec);
  const aset = filterPengusahaanDate_(read_(PENGUSAHAAN_DATA_SHEETS.aset), filters);
  const sdm = read_(PENGUSAHAAN_DATA_SHEETS.sdm);
  const tad = read_(PENGUSAHAAN_DATA_SHEETS.tad);

  const pLatest = latestPenjualan_(penjualan);
  const tarifAgg = aggregateTwoCol_(tarif,'Golongan Tarif','Jumlah Pelanggan');
  const kategoriAgg = aggregateTwoCol_(kategori,'Kategori','Jumlah Pelanggan');
  const kecAgg = aggregateTwoCol_(kec,'Kecamatan','Jumlah Pelanggan');
  const asetAgg = asetSummary_(aset);
  const tadAgg = aggregateTwoCol_(tad,'Jenis TAD','Jumlah');

  return {
    ok:true,
    filters:{year:filters.year||'',month:filters.month||''},
    sheets:{
      penjualan:penjualan,
      tarif:tarif,
      kategori:kategori,
      kec:kec,
      aset:aset,
      sdm:sdm,
      tad:tad
    },
    latest:pLatest,
    cards:{
      pelanggan:pLatest ? numPeng_(pLatest['JUMLAH PELANGGAN']) : 0,
      daya:pLatest ? numPeng_(pLatest['DAYA TERSAMBUNG']) : 0,
      penjualanKwh:pLatest ? numPeng_(pLatest['PENJUALAN (KWH)']) : 0,
      pendapatan:pLatest ? numPeng_(pLatest['RP PENDAPATAN']) : 0,
      growth:pLatest ? numPeng_(pLatest['GROWTH KWH YOY']) : 0,
      natural:pLatest ? numPeng_(pLatest['NATURAL GROWTH']) : 0,
      nonNatural:pLatest ? numPeng_(pLatest['NON NATURAL GROWTH']) : 0
    },
    trends:penjualan.map(r=>({
      bulan:r['BULAN']||'',
      pelanggan:numPeng_(r['JUMLAH PELANGGAN']),
      penjualan:numPeng_(r['PENJUALAN (KWH)']),
      growth:numPeng_(r['GROWTH KWH YOY'])
    })),
    tarif:tarifAgg,
    kategori:kategoriAgg,
    kec:kecAgg,
    aset:asetAgg,
    sdm:sdm,
    tad:tadAgg,
    available:{
      penjualan:!!getSheet_(PENGUSAHAAN_DATA_SHEETS.penjualan),
      tarif:!!getSheet_(PENGUSAHAAN_DATA_SHEETS.tarif),
      kategori:!!getSheet_(PENGUSAHAAN_DATA_SHEETS.kategori),
      kec:!!getSheet_(PENGUSAHAAN_DATA_SHEETS.kec),
      aset:!!getSheet_(PENGUSAHAAN_DATA_SHEETS.aset),
      sdm:!!getSheet_(PENGUSAHAAN_DATA_SHEETS.sdm),
      tad:!!getSheet_(PENGUSAHAAN_DATA_SHEETS.tad)
    }
  };
}

function getSheet_(name){
  return db_().getSheetByName(name);
}

function filterPengusahaanPeriod_(rows,f){
  return rows.filter(r=>{
    const ym=periodFromText_(r['BULAN']||r['Tanggal']||'');
    if(f.year && ym.year && String(ym.year)!==String(f.year)) return false;
    if(f.month && Number(f.month)>0 && ym.month && Number(ym.month)!==Number(f.month)) return false;
    return true;
  });
}

function filterPengusahaanDate_(rows,f){
  return rows.filter(r=>{
    if(!f.year && !f.month)return true;
    const ym=periodFromText_(r['Tanggal']||r['BULAN']||'');
    if(f.year && ym.year && String(ym.year)!==String(f.year))return false;
    if(f.month && Number(f.month)>0 && ym.month && Number(ym.month)!==Number(f.month))return false;
    return true;
  });
}

function periodFromText_(v){
  if(v instanceof Date)return {year:v.getFullYear(),month:v.getMonth()+1};
  const s=String(v||'').trim();
  let m=s.match(/^(\d{4})-(\d{1,2})/);
  if(m)return {year:Number(m[1]),month:Number(m[2])};
  m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if(m)return {year:Number(m[3]),month:Number(m[2])};
  const months={
    januari:1,februari:2,maret:3,april:4,mei:5,juni:6,
    juli:7,agustus:8,september:9,oktober:10,november:11,desember:12
  };
  const parts=s.toLowerCase().split(/\s+/);
  return {year:parts.length>1 && /^\d{4}$/.test(parts[1])?Number(parts[1]):null,month:months[parts[0]]||null};
}

function numPeng_(v){
  if(v===null||v===undefined||v==='')return 0;
  if(typeof v==='number')return isFinite(v)?v:0;
  let s=String(v).trim().replace(/[\\s\\u00a0]/g,'');
  if(!s)return 0;
  let negative=false;
  if(/^\\(.*\\)$/.test(s)){negative=true;s=s.slice(1,-1);}
  s=s.replace(/%/g,'').replace(/^(Rp|IDR)/i,'').replace(/[^0-9,.-]/g,'');
  if(!s)return 0;
  const commas=(s.match(/,/g)||[]).length;
  const dots=(s.match(/\\./g)||[]).length;
  if(commas&&dots){
    // Separator paling kanan diperlakukan sebagai desimal; separator lainnya adalah pemisah ribuan.
    const decimal=s.lastIndexOf(',')>s.lastIndexOf('.')?',':'.';
    const group=decimal===','?'.':',';
    s=s.split(group).join('');
    if(decimal===',')s=s.replace(/,/g,'.');
  }else if(commas||dots){
    const sep=commas?',':'.';
    const parts=s.split(sep);
    if(parts.length>2){
      const last=parts[parts.length-1];
      if(last.length===1||last.length===2){
        s=parts.slice(0,-1).join('')+'.'+last;
      }else if(parts.slice(1).every(part=>part.length===3)){
        s=parts.join('');
      }else{
        s=parts.slice(0,-1).join('')+'.'+last;
      }
    }else{
      const left=parts[0].replace(/^[+-]/,'');
      const right=parts[1]||'';
      if(right.length===3&&left.length>=1&&left.length<=3){
        s=parts.join('');
      }else{
        s=parts[0]+'.'+right;
      }
    }
  }
  const n=Number(s);
  if(!isFinite(n))return 0;
  return negative?-n:n;
}

function latestPenjualan_(rows){
  if(!rows.length)return null;
  return rows[rows.length-1];
}

function aggregateTwoCol_(rows,labelKey,valueKey){
  const map={};
  rows.forEach(r=>{
    const label=String(r[labelKey]||'Tidak diketahui').trim();
    if(!label)return;
    map[label]=(map[label]||0)+numPeng_(r[valueKey]);
  });
  return Object.keys(map).map(k=>({label:k,value:map[k]})).sort((a,b)=>b.value-a.value);
}

function asetSummary_(rows){
  let penyulang=0,keypoint=0,gardu=0,jtm=0,jtr=0;
  rows.forEach(r=>{
    penyulang+=numPeng_(r['Penyulang']);
    keypoint+=numPeng_(r['Keypoint']);
    gardu+=numPeng_(r['Gardu']);
    jtm+=numPeng_(r['JTM Kms']);
    jtr+=numPeng_(r['JTR Kms']);
  });
  return {penyulang,keypoint,gardu,jtm,jtr,rows};
}

function read_(name){
  const sh=db_().getSheetByName(name);
  if(!sh||sh.getLastRow()<2)return [];
  const a=sh.getDataRange().getValues(),h=a.shift();
  return a.filter(r=>r.some(v=>v!==''&&v!==null)).map(r=>{
    const o={};h.forEach((x,i)=>{let v=r[i];if(v instanceof Date)v=Utilities.formatDate(v,CFG.TZ,'yyyy-MM-dd');o[x]=v;});return o;
  });
}

function seedDemo_(){
  const samples={
    'Teknik Jaringan':[
      ['TJ001',new Date(2026,8,2),'EMP-01','Empang','Pemeliharaan','Pemeliharaan jaringan','Selesai','Tim Teknik','Normal','',''],
      ['TJ002',new Date(2026,8,8),'EMP-02','Empang','Gangguan','Penanganan gangguan','Proses','Tim Teknik','Tinggi','',''],
      ['TJ003',new Date(2026,8,15),'EMP-03','Empang','Inspeksi','Inspeksi jaringan','Belum','Tim Teknik','Normal','','']
    ],
    'Transaksi Energi':[
      ['TE001',new Date(2026,8,3),'IDP001','MTR001','EDMI','Pascabayar','Pemeliharaan','Empang','Tim TE','Selesai','','',''],
      ['TE002',new Date(2026,8,10),'IDP002','MTR002','Hexing','Prabayar','Pemeriksaan','Empang','Tim TE','Proses','Perlu tindak lanjut','',''],
      ['TE003',new Date(2026,8,16),'IDP003','MTR003','Itron','Pascabayar','Penggantian','Empang','Tim TE','Belum','','','']
    ],
    'Pelayanan Pelanggan':[
      ['PP001',new Date(2026,8,4),'IDP101','Pasang Baru','Pelanggan 1','Empang','Selesai','Tim YAN','Sesuai',''],
      ['PP002',new Date(2026,8,11),'IDP102','Perubahan Daya','Pelanggan 2','Empang','Proses','Tim YAN','Proses',''],
      ['PP003',new Date(2026,8,17),'IDP103','Pengaduan','Pelanggan 3','Empang','Belum','Tim YAN','Tindak lanjut','']
    ],
    'K3L':[
      ['K301',new Date(2026,8,1),'Safety Morning Call','Kantor ULP','Tim K3L','Selesai','','','',''],
      ['K302',new Date(2026,8,9),'CBD','Lokasi Kerja','Tim K3L','Selesai','Tidak ada temuan','Selesai','',''],
      ['K303',new Date(2026,8,18),'Inspeksi K3','Gudang','Tim K3L','Proses','1 temuan','Tindak lanjut','','']
    ],
    'Pengusahaan':[
      ['PG001',new Date(2026,8,1),'Penjualan Tenaga Listrik',100000,98000,'kWh',98,''],
      ['PG002',new Date(2026,8,1),'Pelanggan',10000,9850,'Pelanggan',98.5,''],
      ['PG003',new Date(2026,8,1),'Penerimaan',1000000000,970000000,'Rupiah',97,'']
    ]
  };
  Object.keys(samples).forEach(n=>{
    const sh=db_().getSheetByName(n);
    if(sh&&sh.getLastRow()===1)sh.getRange(2,1,samples[n].length,samples[n][0].length).setValues(samples[n]);
  });
}