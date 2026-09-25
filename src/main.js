import { createClient } from '@supabase/supabase-js';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.mjs?url';
import './style.css';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

const SB_URL = import.meta.env.VITE_SUPABASE_URL || '';
const SB_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY || '';
const ADMIN_EMAIL = (import.meta.env.VITE_ADMIN_EMAIL || 'chaosslaps@gmail.com').trim();
const supabase = SB_URL && SB_KEY ? createClient(SB_URL, SB_KEY) : null;

const MONTHS = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const FREQ = { weekly:'Setiap minggu', biweekly:'Setiap 2 minggu', monthly:'Setiap bulan', custom:'Custom' };
const state = {
  view:'home', previousView:'home', rt:1, year:new Date().getFullYear(), search:'', agendaSearch:'', agendaCategory:'all',
  rts:[], kks:[], programs:[], schedules:[], sessions:[], transactions:[],
  agendas:[], org:null, session:null, loading:false, selectedProgram:'', selectedSchedule:'',
  selectedScheduleRt:'all', reportMonth:'all', reportYear:new Date().getFullYear(), reportRt:'all',
  reportProgram:'all', kkHistory:null, expandedMonths:[], matrixCollectorNames:{}, adminTab:'dashboard', adminKKRt:''
};
const app = document.querySelector('#app');
const esc = (v='') => String(v).replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
const money = n => new Intl.NumberFormat('id-ID',{style:'currency',currency:'IDR',maximumFractionDigits:0}).format(Number(n||0));
const formatNominal = value => { const digits=String(value??'').replace(/\D/g,''); return digits?Number(digits).toLocaleString('id-ID'):''; };
const parseNominal = value => { const digits=String(value??'').replace(/\D/g,''); return digits?Number(digits):0; };
const bindNominalInput = input => { input.addEventListener('focus',()=>{if(parseNominal(input.value)===0)input.value='';else{input.value=input.value.replace(/\D/g,'');input.select();}});input.addEventListener('input',()=>{input.value=formatNominal(input.value);});input.addEventListener('blur',()=>{input.value=formatNominal(input.value);}); };
const dateID = d => d ? new Intl.DateTimeFormat('id-ID',{day:'2-digit',month:'short',year:'numeric'}).format(new Date(`${d}T00:00:00`)) : '—';
const timeNow = d => d ? new Intl.DateTimeFormat('id-ID',{hour:'2-digit',minute:'2-digit',hour12:false}).format(new Date(d)) : '—';
const toast = (msg,type='ok') => { const el=document.createElement('div'); el.className=`toast ${type}`; el.textContent=msg; document.body.appendChild(el); setTimeout(()=>el.remove(),2800); };
const isConfigured = () => !!supabase;
const requireConfig = () => { if(!supabase){toast('Supabase belum dikonfigurasi. Isi .env terlebih dahulu.','err');return false;}return true; };
const authRequired = () => { if(!state.session){loginModal();return false;}return true; };
const rtName = id => state.rts.find(r=>Number(r.id)===Number(id))?.nama || `RT ${String(id).padStart(2,'0')}`;
const statusLabel = n => Number(n||0)>0 ? money(n) : '—';

function shell(){
  app.innerHTML=`<div class="noise"></div><header class="topbar"><button class="iconbtn backIcon" id="backBtn" title="Kembali">←</button><button class="iconbtn homeIcon" data-nav="home" title="Home">⌂</button><div class="brand"><img src="/assets/pm_logo_transparent.png"><div><b>PEMUDA MAWE</b><span>Guyub • Rukun • Sejahtera</span></div></div><nav>
  <button data-nav="dashboard">Dashboard</button><button data-nav="organisasi">Organisasi</button><button data-nav="agenda">Agenda</button><button data-nav="jimpitan">Jimpitan</button><button data-nav="laporan">Laporan</button><button data-nav="dataKK">Data KK</button><button data-nav="admin">Admin</button></nav><button id="menuBtn" class="iconbtn">☰</button></header><main id="page"></main><footer>PEMUDA MAWE <span>•</span> GUYUB RUKUN SEJAHTERA <small>v10 UI</small></footer>`;
  document.querySelector('.topbar').classList.toggle('insideView',state.view!=='home');
  document.querySelectorAll('[data-nav]').forEach(b=>b.onclick=()=>go(b.dataset.nav));
  document.querySelector('#backBtn').onclick=()=>{const target=state.previousView||'home';state.previousView='home';go(target);};
  document.querySelector('#menuBtn').onclick=()=>document.querySelector('nav').classList.toggle('open');
}
function defaultOrg(){return {id:1,ketua:'Budi Santoso',wakil:'Siti Aminah',sekretaris:'Parno',sekretaris2:'',bendahara:'Agus',bendahara2:'',pembina:'',koordinator:{1:'Karman RT 01',2:'Jono RT 02',3:'Sutrisno RT 03',4:'Wagiman RT 04',5:'Supri RT 05',6:'Marno RT 06',7:'Tukijo HUNTAP'},bidang:{keamanan:'',olahraga:'',seni_budaya:'',sosial:'',humas:'',kerohanian:'',lingkungan:''},visi:'Terwujudnya Pemuda Pemudi Mawe yang Guyub, Rukun, Sejahtera',misi:['Guyub','Rukun','Sejahtera']};}

async function loadAll(){
  if(!supabase)return;
  state.loading=true; render();
  const [rts,org,agenda,programs,schedules]=await Promise.all([
    supabase.from('rt').select('*').order('id'),
    supabase.from('organisasi').select('*').eq('id',1).maybeSingle(),
    supabase.from('agenda').select('*').order('tahun').order('tgl'),
    supabase.from('jimpitan_program').select('*').order('created_at',{ascending:false}),
    supabase.from('jimpitan_schedule').select('*, jimpitan_program(nama_program,jenis,target_nominal)').order('tanggal')
  ]);
  [rts,org,agenda,programs,schedules].forEach(x=>{if(x.error)console.warn(x.error);});
  state.rts=rts.data||[]; state.org=org.data||defaultOrg(); state.agendas=agenda.data||[]; state.programs=programs.data||[]; state.schedules=schedules.data||[];
  state.loading=false; render();
}
async function loadKK(){
  if(!requireConfig())return;
  const {data,error}=await supabase.from('kk').select('*').eq('aktif',true).order('nama');
  if(error){toast(error.message,'err');return;}
  state.kks=data||[];
}
async function loadScheduleTransactions(scheduleId){
  if(!supabase||!scheduleId){state.transactions=[];return;}
  const {data,error}=await supabase.from('jimpitan_transaction').select('*, kk(nama,rt), penarikan_session(nama_penarik,started_at,finished_at)').eq('schedule_id',scheduleId).order('confirmed_at');
  if(error){toast(error.message,'err');return;} state.transactions=data||[];
}
async function init(){
  shell();render();
  if(supabase){
    const {data}=await supabase.auth.getSession(); state.session=data.session;
    supabase.auth.onAuthStateChange((_e,s)=>{state.session=s;});
    await loadAll();
    await loadKK(); await loadReport();
  } else { state.rts=[1,2,3,4,5,6,7].map(i=>({id:i,nama:i===7?'HUNTAP':`RT ${String(i).padStart(2,'0')}`}));state.org=defaultOrg();render(); }
}

function go(view){
  const allowed = ['home','dashboard','organisasi','agenda','jimpitan','dataKK','penarikan','penarikanActive','laporan','admin'];
  if(!allowed.includes(view)) view='home';
  if(view!==state.view)state.previousView=state.view;
  if(view==='admin')state.adminTab='dashboard';
  state.view=view;
  window.scrollTo({top:0,behavior:'instant'});
  render();
}

function render(){
  shell();const p=document.querySelector('#page');
  if(state.loading){p.innerHTML='<div class="loading">Memuat data…</div>';return;}
  const views={home,dashboard,organisasi:dashboard,jimpitan,dataKK,penarikan,penarikanActive,laporan,agenda:agendaPage,admin};
  p.innerHTML=(views[state.view]||home)();bind();
}

function home(){return `<section class="homePage"><div class="homeOrnament top"></div><div class="homeOrnament left"></div><div class="homeOrnament right"></div><div class="homeCard"><div class="homeCrown">✦</div><img class="heroLogo" src="/assets/pm_logo_transparent.png" alt="PM"><div class="heroTitle"><span>PEMUDA, PEMUDI</span><strong>MAWE</strong></div><div class="heroTag"><i></i><span>GUYUB RUKUN SEJAHTERA</span><i></i></div><div class="homeActions"><button class="homeGlow adminGlow" data-nav="admin"><span>♙</span> Admin</button><button class="homeGlow orgGlow" data-nav="organisasi"><span>♧</span> Organisasi</button><button class="homeGlow agendaGlow" data-nav="agenda"><span>▣</span> Agenda</button><button class="homeGlow jimpitanGlow" data-nav="jimpitan"><span>◉</span> Jimpitan</button></div></div></section>`;}
function dashboard(){
 const o=state.org||defaultOrg();
 const bidang=[['Keamanan','keamanan'],['Olahraga','olahraga'],['Seni & Budaya','seni_budaya'],['Sosial','sosial'],['Humas & Informasi','humas'],['Kerohanian','kerohanian'],['Lingkungan','lingkungan']];
 const node=(label,name,cls='')=>`<div class="orgCard ${cls}"><div class="orgAvatar">${label==='KETUA'?'♛':'●'}</div><div><small>${label}</small><b>${esc(name||'Belum diisi')}</b></div></div>`;
 return `<section class="section orgPage"><div class="pageHero"><div><div class="eyebrow">♧ STRUKTUR ORGANISASI</div><h1>STRUKTUR ORGANISASI<br><em>PEMUDA MAWE</em></h1><p>Bersama, Mengabdi, Membangun Desa</p></div><div class="heroQuote">“Pemuda yang solid adalah kekuatan<br>untuk masa depan yang lebih baik”</div></div>
 <div class="orgCanvas"><div class="orgStairNew">${node('PELINDUNG & PENASEHAT',o.pembina,'orgTop')}<span class="connector"></span>${node('KETUA',o.ketua,'orgTop chief')}<span class="connector"></span>${node('WAKIL KETUA',o.wakil,'orgTop')}<span class="connector"></span><div class="orgPair"><div>${node('SEKRETARIS I',o.sekretaris)}${o.sekretaris2?`<div class="orgSub">${node('SEKRETARIS II',o.sekretaris2)}</div>`:''}</div><div>${node('BENDAHARA I',o.bendahara)}${o.bendahara2?`<div class="orgSub">${node('BENDAHARA II',o.bendahara2)}</div>`:''}</div></div><span class="connector"></span><div class="orgLabel">SEKSI / BIDANG</div><div class="orgFields">${bidang.map(([label,key])=>`<div class="fieldGroup">${node(`SIE ${label.toUpperCase()}`,o.bidang?.[key]||'Belum diisi')}<div class="fieldMembers"><span>Anggota 1</span><span>Anggota 2</span></div></div>`).join('')}</div></div></div>
 <div class="orgFooterGrid"><div class="panel orgWilayah"><div class="panelTitle">KOORDINATOR RT / HUNTAP</div><div class="rtGrid">${state.rts.map(r=>`<div class="rtCard"><span>${esc(r.nama)}</span><b>${esc(o.koordinator?.[r.id]||'Belum diisi')}</b></div>`).join('')}</div></div><div class="panel orgStats"><div class="panelTitle">RINGKASAN ORGANISASI</div><div class="miniStats"><div><b>${state.kks.length}</b><span>KK aktif</span></div><div><b>${state.agendas.length}</b><span>Agenda</span></div><div><b>${state.programs.filter(x=>x.aktif!==false).length}</b><span>Program</span></div></div></div></div>
 </section>`;
}
function person(role,name,theme=''){return `<div class="person ${theme}"><small>${role}</small><b>${esc(name)}</b></div>`;}

function programMonthCount(program){
 if(!program?.tanggal_mulai||!program?.tanggal_selesai)return 1;
 const start=new Date(`${program.tanggal_mulai}T00:00:00`),end=new Date(`${program.tanggal_selesai}T00:00:00`);
 return Math.max(1,(end.getFullYear()-start.getFullYear())*12+end.getMonth()-start.getMonth()+1);
}
function programMonths(program){
 if(!program?.tanggal_mulai||!program?.tanggal_selesai)return MONTHS.map((label,month)=>({label,month,year:state.reportYear}));
 const start=new Date(`${program.tanggal_mulai}T00:00:00`),count=programMonthCount(program);
 return Array.from({length:count},(_,index)=>{const date=new Date(start);date.setMonth(start.getMonth()+index);return {label:MONTHS[date.getMonth()],month:date.getMonth(),year:date.getFullYear()};});
}
function jimpitanDisplayTotals(rows){
 const totals=new Map();
 const filled=new Map();
 const payments=[...rows].sort((a,b)=>String(a.jimpitan_schedule?.tanggal||'').localeCompare(String(b.jimpitan_schedule?.tanggal||''))||String(a.confirmed_at||'').localeCompare(String(b.confirmed_at||'')));
 payments.forEach(transaction=>{
  const program=state.programs.find(x=>String(x.id)===String(transaction.jimpitan_schedule?.program_id||transaction.jimpitan_schedule?.jimpitan_program?.id))||transaction.jimpitan_schedule?.jimpitan_program;
  const monthly=Number(program?.target_nominal||0)/programMonthCount(program);
  let remaining=Number(transaction.nominal||0);
  if(!remaining||!monthly)return;
  const date=transaction.jimpitan_schedule?.tanggal;
  let month=date?Number(date.slice(5,7))-1:0;
  let year=date?Number(date.slice(0,4)):state.reportYear;
  let guard=0;
  while(remaining>0.0001&&guard<120){
  const allocationKey=`${transaction.kk_id}:${program?.id||'program'}:${year*12+month}`;
   const used=filled.get(allocationKey)||0;
   const room=Math.max(0,monthly-used);
   const allocation=Math.min(remaining,room);
   if(allocation){
    filled.set(allocationKey,used+allocation);
    const totalKey=`${transaction.kk_id}:${year*12+month}`;
    totals.set(totalKey,(totals.get(totalKey)||0)+allocation);
    remaining-=allocation;
   }
  month++;if(month>11){month=0;year++;}guard++;
  }
 });
 return (kkId,year,month)=>totals.get(`${kkId}:${year*12+month}`)||0;
}

function jimpitan(){
 const rows=state.reportTransactions||[];
 const selectedRt=state.reportRt==='all'?'all':Number(state.reportRt);
 const selectedProgram=state.reportProgram==='all'?null:state.programs.find(p=>String(p.id)===String(state.reportProgram));
 const matrixMonths=programMonths(selectedProgram);
 const visibleKks=state.kks.filter(k=>selectedRt==='all'||Number(k.rt)===selectedRt).filter(k=>!state.search||k.nama.toLowerCase().includes(state.search.toLowerCase()));
 const displayTotal=jimpitanDisplayTotals(rows);
 const monthTotals=(kkId,month,year)=>displayTotal(kkId,year,month);
 const collectorNames=({month,year})=>{if(selectedRt==='all')return 'Pilih RT';const key=`${selectedRt}:${year}:${month+1}`;return state.matrixCollectorNames[key]||(()=>{const names=[...new Set(rows.filter(x=>(Number(x.kk?.rt)===selectedRt)&&String(x.jimpitan_schedule?.tanggal||'').slice(0,7)===`${year}-${String(month+1).padStart(2,'0')}`).map(x=>x.penarikan_session?.nama_penarik).filter(Boolean).filter(x=>String(x).trim().toLowerCase()!=='admin'))];return names.length?names.join(', '):'—';})();};
 const grand=month=>visibleKks.reduce((a,k)=>a+monthTotals(k.id,month.month,month.year),0);
 return `<section class="section wide jimpitanPage"><div class="pageHero compactHero"><div><div class="eyebrow">◉ LAPORAN JIMPITAN</div><h1>JIMPITAN</h1><p>Bersama, Ringan di Tangan, Berat di Keikhlasan</p></div><div class="heroActionGroup">${state.session?`<button class="primary" id="newProgram">+ PROGRAM BARU</button><button class="secondary" data-nav="admin">KELOLA PROGRAM</button>`:`<button class="secondary" data-nav="laporan">LIHAT LAPORAN</button>`}</div></div>
 <div class="jimpitanToolbar"><div class="rtPills">${state.rts.map(r=>`<button class="rtPill ${String(state.reportRt)===String(r.id)?'active':''}" data-report-rt="${r.id}">${esc(r.nama)}</button>`).join('')}<button class="rtPill ${state.reportRt==='all'?'active':''}" data-report-rt="all">Semua</button></div><div class="jimpitanTools"><select id="reportProgram"><option value="all">Semua Jenis / Program</option>${state.programs.map(p=>`<option value="${p.id}" ${String(p.id)===String(state.reportProgram)?'selected':''}>${esc(p.nama_program)}</option>`).join('')}</select>${state.session?`<button class="secondary" id="addKK">+ KK</button><button class="secondary" id="importKK">Upload KK</button>`:''}</div></div>
 <div class="jimpitanMatrixPanel"><div class="matrixCaption"><div><b>${esc(selectedProgram?.nama_program||'Semua Program Jimpitan')}</b><span>${selectedProgram?`Target ${money(selectedProgram.target_nominal)} / KK • Rata-rata ${money(selectedProgram.target_nominal/programMonthCount(selectedProgram))} / bulan`:'Rata-rata cicilan per bulan'}</span></div><span>${selectedRt==='all'?'Semua RT':esc(rtName(selectedRt))} • ${state.reportYear}</span></div><div class="tableWrap matrixDarkWrap"><table class="jimpitanMatrix darkMatrix"><thead><tr><th class="sticky">No</th><th class="sticky second">Nama KK</th><th>Iuran</th>${matrixMonths.map(m=>`<th>${m.label.slice(0,3).toUpperCase()}<small>${m.year}</small></th>`).join('')}<th>TOTAL</th></tr></thead><tbody>${visibleKks.length?visibleKks.map((k,i)=>{const vals=matrixMonths.map(m=>monthTotals(k.id,m.month,m.year));const total=vals.reduce((a,b)=>a+b,0);return `<tr><td class="sticky rowNo">${String(i+1).padStart(2,'0')}</td><td class="sticky second"><b>${esc(k.nama)}</b><small>${esc(rtName(k.rt))}</small></td><td>${selectedProgram?money(selectedProgram.target_nominal):'—'}</td>${vals.map(v=>`<td><span class="monthCell">${v?money(v):'Rp 0'}</span></td>`).join('')}<td><b>${total?money(total):'Rp 0'}</b></td></tr>`}).join(''):`<tr><td colspan="${matrixMonths.length+4}" class="emptyCell">Belum ada data KK.</td></tr>`}</tbody><tfoot><tr><th class="sticky">—</th><th class="sticky second">TOTAL</th><th>—</th>${matrixMonths.map(month=>`<th>${grand(month)?money(grand(month)):'Rp 0'}</th>`).join('')}<th>${money(matrixMonths.reduce((a,month)=>a+grand(month),0))}</th></tr><tr class="collectorRow"><th class="sticky">—</th><th class="sticky second">PENARIK</th><th>—</th>${matrixMonths.map(month=>`<th><div class="collectorNames">${esc(collectorNames(month))}</div></th>`).join('')}<th>—</th></tr></tfoot></table></div></div>
 <div class="jimpitanBottom"><div class="panel"><div class="panelTitle">JENIS JIMPITAN</div><div class="typeChips">${[...new Set(state.programs.map(p=>p.jenis).filter(Boolean))].map(x=>`<span>${esc(x)}</span>`).join('')||'<span class="emptyChip">Belum ada jenis</span>'}</div></div><div class="panel jimpitanHelp"><div class="panelTitle">ALUR PENARIKAN</div><p>Jadwal dibuat otomatis dari Program Jimpitan. Penarik mengisi identitas sendiri saat mulai penarikan, lalu setiap KK tetap tercatat termasuk nominal Rp0.</p><button class="secondary" data-nav="penarikan">BUKA PENARIKAN</button></div></div></section>`;
}
function programCard(p){return `<article class="programCard"><div class="programTop"><span class="badge">${esc(p.jenis)}</span><span class="badge ${p.aktif?'ok':''}">${p.aktif?'AKTIF':'NONAKTIF'}</span></div><h3>${esc(p.nama_program)}</h3><div class="programMeta"><span>Target</span><b>${money(p.target_nominal)}</b><span>Periode</span><b>${dateID(p.tanggal_mulai)} — ${dateID(p.tanggal_selesai)}</b><span>Frekuensi</span><b>${esc(FREQ[p.frekuensi]||p.frekuensi)}</b></div><div class="cardActions">${state.session?`<button class="secondary small" data-edit-program="${p.id}">Edit</button><button class="dangerText" data-toggle-program="${p.id}">${p.aktif?'Nonaktifkan':'Aktifkan'}</button>`:''}</div></article>`;}
function scheduleList(){
 let arr=state.schedules.filter(s=>!state.selectedProgram||s.program_id===state.selectedProgram);
 if(!arr.length)return '<div class="empty">Belum ada jadwal.</div>';
 return arr.map(s=>`<div class="scheduleRow"><div class="scheduleDate"><b>${new Date(`${s.tanggal}T00:00:00`).getDate()}</b><span>${MONTHS[new Date(`${s.tanggal}T00:00:00`).getMonth()].slice(0,3)}</span></div><div class="scheduleInfo"><b>${esc(s.jimpitan_program?.nama_program||'Program')}</b><span>${esc(s.jimpitan_program?.jenis||'')} • ${esc(rtName(s.rt))}</span><small>${s.status==='finished'?'Selesai':s.status==='active'?'Sedang berjalan':'Terjadwal'}</small></div><button class="primary small" data-start-schedule="${s.id}">${s.status==='active'?'Lanjut':'Mulai'}</button></div>`).join('');
}

function dataKK(){
 const grouped=state.rts.map(r=>({...r,kks:state.kks.filter(k=>Number(k.rt)===Number(r.id)&&(!state.search||k.nama.toLowerCase().includes(state.search.toLowerCase())))}));
 return `<section class="section wide"><div class="sectionHead"><div><div class="eyebrow">DATA KK</div><h2>Data KK per RT</h2><p class="muted">Urutan otomatis A–Z. KK tidak hilang hanya karena nominal penarikan Rp0.</p></div>${state.session?`<div class="cardActions"><button class="primary" id="addKK">+ Tambah KK</button><button class="secondary" id="importKK">Import PDF / Excel / CSV / TXT</button></div>`:''}</div>
 <div class="toolbar"><select id="kkRt"><option value="all">Semua wilayah</option>${state.rts.map(r=>`<option value="${r.id}" ${String(state.rt)===String(r.id)?'selected':''}>${esc(r.nama)}</option>`).join('')}</select><input id="kkSearch" value="${esc(state.search)}" placeholder="Cari nama KK…"></div>
 <div class="rtDataGrid">${grouped.map(g=>`<article class="panel"><div class="panelTitle">${esc(g.nama)} <span class="count">${g.kks.length}</span></div><ol class="kkList">${g.kks.map(k=>`<li><b>${esc(k.nama)}</b>${state.session?`<span><button class="linkBtn" data-history-kk="${k.id}">Riwayat</button><button class="dangerText" data-delete-kk="${k.id}">Nonaktifkan</button></span>`:''}</li>`).join('')||'<li class="muted">Belum ada KK.</li>'}</ol></article>`).join('')}</div></section>`;
}

function penarikan(){
 const schedules=state.schedules.filter(s=>s.status!=='finished').slice(0,30);
 return `<section class="section wide"><div class="sectionHead"><div><div class="eyebrow">PENARIKAN</div><h2>Jadwal Penarikan</h2><p class="muted">Pilih jadwal, lalu penarik mengisi identitasnya sendiri. Tidak perlu ditentukan Admin.</p></div></div>
 <div class="scheduleGrid">${schedules.length?schedules.map(s=>`<div class="scheduleRow"><div class="scheduleDate"><b>${new Date(`${s.tanggal}T00:00:00`).getDate()}</b><span>${MONTHS[new Date(`${s.tanggal}T00:00:00`).getMonth()].slice(0,3)}</span></div><div class="scheduleInfo"><b>${esc(s.jimpitan_program?.nama_program||'')}</b><span>${esc(s.jimpitan_program?.jenis||'')} • ${esc(rtName(s.rt))}</span><small>Target: ${money(s.jimpitan_program?.target_nominal||0)}</small></div><button class="primary small" data-start-schedule="${s.id}">Mulai Penarikan</button></div>`).join(''):'<div class="empty">Tidak ada jadwal yang tersedia.</div>'}</div>
 ${state.kkHistory?historyPanel():''}</section>`;
}

function historyPanel(){
 const k=state.kkHistory; return `<div class="modalLike panel"><div class="sectionHead"><div><div class="panelTitle">Riwayat KK</div><h3>${esc(k.nama)}</h3></div><button class="secondary small" id="closeHistory">Tutup</button></div><div class="historyList">${(k.rows||[]).map(t=>`<div class="historyRow"><div><b>${dateID(t.jimpitan_schedule?.tanggal)}</b><span>${esc(t.jimpitan_schedule?.jimpitan_program?.nama_program||'')} • ${esc(rtName(t.kk?.rt))}</span></div><div><b>${statusLabel(t.nominal)}</b><span>Penarik: ${esc(t.penarikan_session?.nama_penarik||'—')}</span><small>${t.confirmed_at?dateID(t.confirmed_at.slice(0,10))+' • '+timeNow(t.confirmed_at):'Belum dikonfirmasi'}</small></div>${state.session?`<button class="secondary small" data-correct="${t.id}">Koreksi</button>`:''}</div>`).join('')||'<div class="empty">Belum ada transaksi.</div>'}</div></div>`;
}

async function openSchedule(id){
 const s=state.schedules.find(x=>String(x.id)===String(id)); if(!s)return;
 const {data:existing}=await supabase.from('penarikan_session').select('*').eq('schedule_id',id).order('started_at',{ascending:false}).limit(1);
 const session=existing?.[0];
 if(session){state.selectedSchedule=id;state.sessions=[session];await loadScheduleTransactions(id);go('penarikanActive');return;}
 modal(`
  <div class="modalCard compact">
   <div class="modalHead"><div><div class="eyebrow">MULAI PENARIKAN</div><h3>${esc(rtName(s.rt))} — ${esc(dateID(s.tanggal))}</h3></div><button class="modalClose" data-close-modal>×</button></div>
   <label>Nama Penarik<input id="collectorName" autocomplete="name" placeholder="Masukkan nama Anda"></label>
   <label>Identitas tambahan <span class="muted">(opsional)</span><input id="collectorIdentity" autocomplete="off" placeholder="Contoh: nomor HP"></label>
   <div class="modalActions"><button class="secondary" data-close-modal>Batal</button><button class="primary" id="startCollectionModal">Mulai Penarikan</button></div>
  </div>`, 'collectorName');
 document.querySelector('#startCollectionModal').onclick=async()=>{
  const name=document.querySelector('#collectorName').value.trim();
  const identity=document.querySelector('#collectorIdentity').value.trim();
  if(!name)return toast('Nama penarik wajib diisi','err');
  const {data,error}=await supabase.from('penarikan_session').insert({schedule_id:id,nama_penarik:name,identitas:identity,started_at:new Date().toISOString(),status:'active'}).select().single();
  if(error){toast(error.message,'err');return;}
  closeModal();
  await supabase.from('jimpitan_schedule').update({status:'active'}).eq('id',id);
  state.selectedSchedule=id;state.sessions=[data];await loadAll();await loadScheduleTransactions(id);go('penarikanActive');
 };
 document.querySelector('#collectorIdentity').onkeydown=e=>{if(e.key==='Enter')document.querySelector('#startCollectionModal').click();};
}
function penarikanActive(){
 const s=state.schedules.find(x=>String(x.id)===String(state.selectedSchedule));const session=state.sessions[0];if(!s||!session)return `<div class="empty">Jadwal tidak ditemukan.</div>`;
 const program=s.jimpitan_program||state.programs.find(p=>p.id===s.program_id);
 const kk=state.kks.filter(k=>Number(k.rt)===Number(s.rt));
 return `<section class="section wide"><div class="sectionHead"><div><div class="eyebrow">PENARIKAN AKTIF</div><h2>${esc(program?.nama_program||'Jimpitan')}</h2><p class="muted">${dateID(s.tanggal)} • ${esc(rtName(s.rt))}</p></div><div class="collectorBox"><small>Penarik</small><b>${esc(session.nama_penarik)}</b><span>Mulai ${timeNow(session.started_at)}</span></div></div>
 <div class="notice">Setiap KK tetap tampil. Jika tidak memberi uang, masukkan <b>0</b> atau kosong lalu konfirmasi. Nilai tersimpan sebagai Rp0.</div>
 <div class="collectionGrid">${kk.map(k=>transactionCard(k,s,session)).join('')||'<div class="empty">Belum ada KK untuk wilayah ini.</div>'}</div>
 <div class="sectionActions"><button class="secondary" id="backPenarikan">Kembali</button><button class="primary" id="finishCollection">Selesai Penarikan</button></div></section>`;
}
function transactionCard(k,s,session){
 const t=state.transactions.find(x=>x.kk_id===k.id);
 const done=!!t;
 return `<div class="collectionCard ${done?'done':''}"><div class="collectionName"><b>${esc(k.nama)}</b><span>${esc(rtName(k.rt))}</span></div><div class="moneyInput"><input type="text" inputmode="numeric" value="${done?formatNominal(t.nominal):''}" placeholder="0" data-nominal="${k.id}" ${done?'disabled':''}><span>Rp</span></div><button class="confirmBtn" data-confirm-kk="${k.id}" ${done?'disabled':''}>${done?'✓ Tercatat':'✓ Konfirmasi'}</button>${done?`<div class="confirmed">✓ ${statusLabel(t.nominal)} • ${timeNow(t.confirmed_at)}</div>`:''}</div>`;
}

function laporan(){
 const rows=state.reportTransactions||[];
 const total=rows.reduce((a,x)=>a+Number(x.nominal||0),0);
 const selectedProgram=state.reportProgram==='all'?null:state.programs.find(p=>String(p.id)===String(state.reportProgram));
 const selectedRt=state.reportRt==='all'?'all':Number(state.reportRt);
 const visibleKks=state.kks.filter(k=>selectedRt==='all'||Number(k.rt)===selectedRt).filter(k=>!state.search||k.nama.toLowerCase().includes(state.search.toLowerCase()));
 const monthTotals=(kkId,month)=>rows.filter(x=>x.kk_id===kkId && x.jimpitan_schedule?.tanggal?.slice(5,7)===String(month).padStart(2,'0')).reduce((a,x)=>a+Number(x.nominal||0),0);
 const grandMonth=m=>visibleKks.reduce((sum,k)=>sum+monthTotals(k.id,m),0);
 const collectorNames=m=>{
  const names=[...new Set(rows.filter(x=>x.jimpitan_schedule?.tanggal?.slice(5,7)===String(m).padStart(2,'0')).map(x=>x.penarikan_session?.nama_penarik).filter(Boolean).map(x=>String(x).trim()).filter(Boolean).filter(x=>x.toLowerCase()!=='admin'))];
   return names.length?names.join(', '):'—';
 };
 const target=selectedProgram?.target_nominal||0;
 return `<section class="section wide reportPage"><div class="reportTop"><button class="iconbtn" data-nav="home">←</button><div><div class="eyebrow">LAPORAN JIMPITAN</div><h2>JIMPITAN ${state.rts.length} RT</h2></div><div class="reportTopActions"><button class="secondary small" id="exportExcel">EXPORT EXCEL</button><button class="primary small" id="exportPdf">EXPORT PDF</button></div></div>
 <div class="reportFilters"><div class="pillGroup">${state.rts.map(r=>`<button class="filterPill ${String(state.reportRt)===String(r.id)?'active':''}" data-report-rt="${r.id}">${esc(r.nama)}</button>`).join('')}<button class="filterPill ${state.reportRt==='all'?'active':''}" data-report-rt="all">SEMUA</button></div><select id="reportProgram"><option value="all">Semua Jenis / Program</option>${state.programs.map(p=>`<option value="${p.id}" ${String(p.id)===String(state.reportProgram)?'selected':''}>${esc(p.nama_program)}</option>`).join('')}</select><select id="reportYear">${[2026,2027,2028].map(y=>`<option value="${y}" ${y===state.reportYear?'selected':''}>${y}</option>`).join('')}</select><select id="reportMonth"><option value="all">Semua Bulan</option>${MONTHS.map((m,i)=>`<option value="${i+1}" ${String(i+1)===String(state.reportMonth)?'selected':''}>${m}</option>`).join('')}</select></div>
 <div class="reportTitleRow"><div><b>${esc(selectedProgram?.nama_program||'Semua Program Jimpitan')}</b><span>${selectedProgram?`Target per KK: ${money(target)}`:'Rekap nominal aktual per bulan'}</span></div><div><b>${selectedRt==='all'?'Semua Wilayah':esc(rtName(selectedRt))}</b><span>Tahun ${state.reportYear}</span></div></div>
 <div class="tableWrap reportTableWrap"><table class="jimpitanMatrix"><thead><tr><th class="sticky">NAMA KK</th><th>RT</th><th>TARGET</th>${MONTHS.map(m=>`<th>${m.slice(0,3).toUpperCase()}</th>`).join('')}<th>TOTAL</th></tr></thead><tbody>${visibleKks.length?visibleKks.map(k=>{const vals=MONTHS.map((_,i)=>monthTotals(k.id,i+1));const sum=vals.reduce((a,b)=>a+b,0);return `<tr><td class="sticky"><b>${esc(k.nama)}</b></td><td>${esc(rtName(k.rt))}</td><td>${selectedProgram?money(target):'—'}</td>${vals.map(v=>`<td>${v?money(v):'—'}</td>`).join('')}<td><b>${sum?money(sum):'—'}</b></td></tr>`}).join(''):`<tr><td colspan="16" class="emptyCell">Belum ada data KK.</td></tr>`}</tbody><tfoot><tr><th class="sticky">TOTAL</th><th>—</th><th>—</th>${MONTHS.map((_,i)=>`<th>${grandMonth(i+1)?money(grandMonth(i+1)):'—'}</th>`).join('')}<th>${total?money(total):'—'}</th></tr><tr class="collectorRow"><th class="sticky">PENARIK</th><th>—</th><th>—</th>${MONTHS.map((_,i)=>`<th><div class="collectorNames">${esc(collectorNames(i+1))}</div></th>`).join('')}<th>—</th></tr></tfoot></table></div>
 <div class="reportSummary"><div><small>Total transaksi</small><b>${rows.length}</b></div><div><small>Total terkumpul</small><b>${money(total)}</b></div><div><small>KK tercatat</small><b>${visibleKks.length}</b></div><div><small>Rp0 / belum menerima</small><b>${visibleKks.filter(k=>!rows.some(x=>x.kk_id===k.id)).length}</b></div></div>
 <details class="reportDetail"><summary>Detail transaksi</summary><div class="tableWrap"><table><thead><tr><th>Tanggal</th><th>RT</th><th>KK</th><th>Jenis</th><th>Program</th><th>Penarik</th><th>Nominal</th><th>Jam</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${dateID(x.jimpitan_schedule?.tanggal)}</td><td>${esc(rtName(x.kk?.rt))}</td><td>${esc(x.kk?.nama||'—')}</td><td>${esc(x.jimpitan_schedule?.jimpitan_program?.jenis||'')}</td><td>${esc(x.jimpitan_schedule?.jimpitan_program?.nama_program||'')}</td><td>${esc(x.penarikan_session?.nama_penarik||'—')}</td><td><b>${statusLabel(x.nominal)}</b></td><td>${timeNow(x.confirmed_at)}</td></tr>`).join('')||'<tr><td colspan="8" class="emptyCell">Belum ada transaksi.</td></tr>'}</tbody></table></div></details>
 </section>`;
}

function agendaPage(){
 const sorted=[...state.agendas].sort((a,b)=>{const da=new Date(`${a.tahun}-${String(MONTHS.indexOf(a.bulan)+1).padStart(2,'0')}-${String(a.tgl).padStart(2,'0')}`);const db=new Date(`${b.tahun}-${String(MONTHS.indexOf(b.bulan)+1).padStart(2,'0')}-${String(b.tgl).padStart(2,'0')}`);return da-db;});
 const category=state.agendaCategory||'all',query=(state.agendaSearch||'').trim().toLowerCase();
 const visible=sorted.filter(a=>{const text=`${a.acara||''} ${a.deskripsi||''} ${a.tempat||''} ${a.bulan||''}`.toLowerCase();return (category==='all'||text.includes(category))&&(!query||text.includes(query));});
 const d=new Date(); const month=d.getMonth(),year=d.getFullYear(),first=new Date(year,month,1),days=new Date(year,month+1,0).getDate(),offset=(first.getDay()+6)%7; const eventDays=new Set(sorted.filter(a=>a.tahun===year&&MONTHS.indexOf(a.bulan)===month).map(a=>Number(a.tgl)));
 const calCells=Array.from({length:offset+days},(_,i)=>i<offset?'':i-offset+1).map(v=>v?`<span class="calDay ${eventDays.has(Number(v))?'hasEvent':''} ${Number(v)===d.getDate()?'today':''}">${v}</span>`:'<span></span>').join('');
 return `<section class="section wide agendaPage darkPage"><div class="pageHero agendaHero"><div><div class="eyebrow">▣ AGENDA PEMUDA MAWE</div><h1>AGENDA<br><em>KEGIATAN</em></h1><p>Jadwal kegiatan, acara, dan agenda penting Pemuda Mawe.</p></div><div class="heroQuote">“Kegiatan hari ini,<br>untuk masa depan<br>yang lebih baik”</div></div><div class="agendaToolbar"><div class="agendaFilters"><button class="filterPill active">Semua</button><button class="filterPill">Rutin</button><button class="filterPill">Sosial</button><button class="filterPill">Kegiatan</button><button class="filterPill">Rapat</button><button class="filterPill">Lainnya</button></div><div class="agendaTools"><input placeholder="Cari agenda..." disabled>${state.session?`<button class="primary" id="newAgenda">+ Tambah Agenda</button>`:''}</div></div><div class="agendaLayout"><div><div class="agendaList">${sorted.length?sorted.map(a=>`<article class="agendaCard"><div class="dateBox"><b>${esc(a.tgl)}</b><span>${esc((a.bulan||'').slice(0,3))}</span><small>${esc(String(a.tahun))}</small></div><div class="agendaBody"><div class="agendaTop"><h3>${esc(a.acara)}</h3>${state.session?`<button class="dangerText" data-del-agenda="${a.id}">×</button>`:''}</div><p>◷ ${esc(a.jam)}</p><p>⌖ ${esc(a.tempat)}</p><div class="desc">${esc(a.deskripsi||'')}</div><div class="cardActions">${state.session?`<button class="secondary small" data-edit-agenda="${a.id}">EDIT</button>`:''}<button class="wa small" data-wa-agenda="${a.id}">◉ SHARE KE WA</button></div></div></article>`).join(''):`<div class="empty">Belum ada agenda.</div>`}</div></div><aside class="agendaAside"><div class="panel miniCalendar"><div class="calendarHead"><b>${MONTHS[month]} ${year}</b><span>‹　›</span></div><div class="weekHead">${['Sen','Sel','Rab','Kam','Jum','Sab','Min'].map(x=>`<b>${x}</b>`).join('')}</div><div class="calendarGrid">${calCells}</div></div><div class="panel upcoming"><div class="panelTitle">AGENDA MENDATANG</div>${sorted.slice(0,5).map(a=>`<div class="upcomingRow"><strong>${String(a.tgl).padStart(2,'0')}<small>${(a.bulan||'').slice(0,3)}</small></strong><span>${esc(a.acara)}<small>${esc(a.jam)}</small></span></div>`).join('')||'<div class="empty">Belum ada agenda.</div>'}</div><div class="panel quickPanel"><div class="panelTitle">QUICK ACTIONS</div><div class="quickGrid"><button data-new-agenda>+ Tambah Agenda</button><button data-nav="agenda">Lihat Semua</button><button id="shareAllAgenda">Share Semua</button><button id="exportPdf">Export PDF</button></div></div></aside></div></section>`;
}
function admin(){
 if(!state.session)return `<section class="section adminLoginPage"><div class="loginHero"><div class="eyebrow">♙ PANEL ADMIN</div><h1>MASUK KE<br><em>ADMIN</em></h1><p>Kelola data, pantau kegiatan, dan bangun kebersamaan.</p><button class="primary loginBig" id="login">Masuk sebagai Admin</button></div></section>`;
 const total=state.kks.length,agenda=state.agendas.length,program=state.programs.filter(x=>x.aktif!==false).length;
 return `<section class="section wide adminPage"><div class="adminHero"><div><div class="eyebrow">♙ PANEL ADMIN</div><h1>Selamat Datang,<br><em>ADMIN</em></h1><p>Kelola data, pantau kegiatan, bangun kebersamaan.</p></div><img src="/assets/pm_logo_transparent.png" alt="PM"><div class="adminUser"><span>● Online</span><b>Admin</b><small>Super Admin</small><button id="logout" class="ghostBtn">Keluar</button></div></div><div class="adminStats"><div class="adminStat"><span>♧</span><small>Total Anggota</small><b>${total}</b><em>Data KK aktif</em></div><div class="adminStat"><span>▣</span><small>Agenda Aktif</small><b>${agenda}</b><em>Kegiatan tersimpan</em></div><div class="adminStat"><span>◉</span><small>Total Program</small><b>${program}</b><em>Program jimpitan</em></div><div class="adminStat"><span>▰</span><small>Total Jimpitan</small><b>${money((state.reportTransactions||[]).reduce((a,x)=>a+Number(x.nominal||0),0))}</b><em>Rekap transaksi</em></div></div><div class="adminDashboardGrid"><div>${adminPanel()}</div><aside class="adminSide"><div class="panel quickAdmin"><div class="panelTitle">AKSI CEPAT</div><div class="quickAdminGrid"><button data-add-kk>＋ Tambah KK</button><button data-new-agenda>▣ Tambah Agenda</button><button data-new-program>◉ Tambah Jimpitan</button><button data-nav="dataKK">▤ Data KK</button><button data-nav="laporan">▤ Lihat Laporan</button><button data-admin-tab="struktur">⚙ Pengaturan</button></div></div><div class="panel adminQuote"><div class="quoteMark">“</div><p>“Pemuda yang solid adalah kekuatan untuk masa depan yang lebih baik.”</p><small>— Pemuda Mawe —</small></div></aside></div></section>`;
}
function adminPanel(){
 const o=state.org||defaultOrg(); const bidang=[['Keamanan','keamanan'],['Olahraga','olahraga'],['Seni & Budaya','seni_budaya'],['Sosial','sosial'],['Humas & Informasi','humas'],['Kerohanian','kerohanian'],['Lingkungan','lingkungan']];
 return `<div class="adminManage"><div class="adminTabs"><button class="tab active" data-admin-tab="jimpitan">Jimpitan</button><button class="tab" data-admin-tab="agenda">Agenda</button><button class="tab" data-admin-tab="kk">Data KK</button><button class="tab" data-admin-tab="struktur">Organisasi</button></div><div class="adminSection" id="adminJimpitan"><div class="panel"><div class="panelTitle">KELOLA JIMPITAN</div><p class="muted">Atur program, jenis, target, periode, frekuensi dan wilayah. Jadwal penarikan dibuat otomatis sesuai program.</p><button class="primary" id="newProgram">+ Buat Program</button><div class="adminList">${state.programs.map(p=>`<div class="adminItem"><div><b>${esc(p.nama_program)}</b><span>${esc(p.jenis)} • ${money(p.target_nominal)} • ${esc(FREQ[p.frekuensi]||p.frekuensi)}</span></div><div><button class="secondary small" data-edit-program="${p.id}">Edit</button><button class="dangerText" data-toggle-program="${p.id}">${p.aktif?'Nonaktifkan':'Aktifkan'}</button></div></div>`).join('')||'<div class="muted">Belum ada program.</div>'}</div></div></div><div class="adminSection" id="adminAgenda"><div class="panel"><div class="panelTitle">KELOLA AGENDA</div><p class="muted">Tambah, edit, dan hapus agenda kegiatan Pemuda Mawe.</p><button class="primary" id="newAgenda">+ Tambah Agenda</button><div class="adminList">${state.agendas.length?state.agendas.slice().sort((a,b)=>{const da=new Date(`${a.tahun}-${String(MONTHS.indexOf(a.bulan)+1).padStart(2,'0')}-${String(a.tgl).padStart(2,'0')}`);const db=new Date(`${b.tahun}-${String(MONTHS.indexOf(b.bulan)+1).padStart(2,'0')}-${String(b.tgl).padStart(2,'0')}`);return da-db;}).map(a=>`<div class="adminItem"><div><b>${esc(a.acara)}</b><span>${esc(a.tgl)} ${esc(a.bulan)} ${esc(a.tahun)} • ${esc(a.tempat)}</span></div><div><button class="secondary small" data-edit-agenda="${a.id}">Edit</button><button class="dangerText" data-del-agenda="${a.id}">Hapus</button></div></div>`).join(''):'<div class="muted">Belum ada agenda.</div>'}</div></div></div><div class="adminSection" id="adminKK"><div class="panel"><div class="panelTitle">DATA KK</div><p class="muted">Tambah manual atau import PDF, Excel, dan CSV. Nama otomatis A–Z dan riwayat tetap aman.</p><div class="cardActions"><button class="primary" data-add-kk>+ Tambah KK</button><button class="secondary" id="importKK">Import PDF / Excel / CSV</button></div></div><div class="panel"><div class="panelTitle">JENIS PROGRAM</div><div class="typeChips">${[...new Set(state.programs.map(p=>p.jenis).filter(Boolean))].map(x=>`<span>${esc(x)}</span>`).join('')||'<span class="emptyChip">Belum ada jenis program</span>'}</div></div></div><div class="adminSection" id="adminStruktur"><div class="panel"><div class="panelTitle">STRUKTUR ORGANISASI KEPEMUDAAN</div><div class="twoCol">${[['pembina','Pelindung & Penasehat'],['ketua','Ketua'],['wakil','Wakil Ketua'],['sekretaris','Sekretaris I'],['sekretaris2','Sekretaris II'],['bendahara','Bendahara I'],['bendahara2','Bendahara II']].map(([f,l])=>`<label>${l}<input id="org_${f}" value="${esc(o[f]||'')}" placeholder="Nama pengurus"></label>`).join('')}</div><div class="panelTitle top">KOORDINATOR BIDANG</div><div class="twoCol">${bidang.map(([l,f])=>`<label>${l}<input id="org_bidang_${f}" value="${esc(o.bidang?.[f]||'')}" placeholder="Nama koordinator"></label>`).join('')}</div><div class="panelTitle top">KOORDINATOR WILAYAH</div><div class="twoCol">${state.rts.map(r=>`<label>${esc(r.nama)}<input id="org_rt_${r.id}" value="${esc(o.koordinator?.[r.id]||'')}"></label>`).join('')}</div><button class="primary" id="saveOrg">Simpan Struktur</button></div></div></div>`;
}
async function saveMatrixName(id,input){
 if(!authRequired())return;
 const name=input.value.trim();if(!name)return toast('Nama KK wajib diisi','err');
 const {error}=await supabase.from('kk').update({nama:name}).eq('id',id);
 if(error){toast(error.message,'err');return;}
 toast('Nama KK tersimpan');await loadKK();await loadReport();render();
}
function collectorNameModal(){
 return new Promise(resolve=>{
  modal(`<div class="modalCard compact"><div class="modalHead"><div><div class="eyebrow">PENARIKAN</div><h3>Nama Penarik</h3><p class="muted">Isi nama Anda sebelum menyimpan nominal bulan ini.</p></div><button class="modalClose" data-close-modal>×</button></div><label>Nama penarik<input id="matrixCollectorName" autocomplete="name" placeholder="Masukkan nama Anda"></label><div class="modalActions"><button class="secondary" data-close-modal>Batal</button><button class="primary" id="saveMatrixCollector">✓ Lanjutkan</button></div></div>`, 'matrixCollectorName');
  const finish=()=>{const name=document.querySelector('#matrixCollectorName')?.value.trim();if(!name)return toast('Nama penarik wajib diisi','err');closeModal();resolve(name);};
  document.querySelector('#saveMatrixCollector').onclick=finish;document.querySelector('#matrixCollectorName').onkeydown=e=>{if(e.key==='Enter')finish();};
  document.querySelector('#appModal [data-close-modal]')?.addEventListener('click',()=>resolve(null),{once:true});
 });
}
async function saveMatrixCollectorName(month,year,input){
 if(!authRequired())return;
 const name=input.value.trim();if(!name)return toast('Nama penarik wajib diisi','err');
 const selectedRt=state.reportRt==='all'?'all':Number(state.reportRt);if(selectedRt==='all')return toast('Pilih RT terlebih dahulu agar nama penarik tidak tercampur','err');
 const collectorKey=`${selectedRt}:${year}:${month+1}`;state.matrixCollectorNames[collectorKey]=name;
 const program=state.reportProgram==='all'?null:state.programs.find(x=>String(x.id)===String(state.reportProgram));
 if(program){
  const selectedRt=state.reportRt==='all'?null:Number(state.reportRt);
  const schedules=state.schedules.filter(x=>x.program_id===program.id&&String(x.tanggal).slice(0,4)===String(year)&&Number(String(x.tanggal).slice(5,7))===month+1&&(!selectedRt||Number(x.rt)===selectedRt));
  const ids=schedules.map(x=>x.id);
  if(ids.length){const {data:sessions,error}=await supabase.from('penarikan_session').select('id').in('schedule_id',ids);if(error){toast(error.message,'err');return;}for(const session of sessions||[]){const {error:updateError}=await supabase.from('penarikan_session').update({nama_penarik:name,identitas:'Input tabel'}).eq('id',session.id);if(updateError){toast(updateError.message,'err');return;}}}
 }
 toast('Nama penarik tersimpan');render();
}
async function saveMatrixNominal(kkId,year,month,input){
 if(!authRequired())return;
 const kk=state.kks.find(x=>String(x.id)===String(kkId));
 const selectedProgram=state.reportProgram==='all'?null:state.programs.find(x=>String(x.id)===String(state.reportProgram));
 const matchingPrograms=state.programs.filter(candidate=>state.schedules.some(schedule=>schedule.program_id===candidate.id&&Number(schedule.rt)===Number(kk?.rt)&&String(schedule.tanggal).slice(0,4)===String(year)&&Number(String(schedule.tanggal).slice(5,7))===month+1));
 const program=selectedProgram||matchingPrograms[0];
 if(!kk||!program)return toast('Pilih program jimpitan terlebih dahulu','err');
 const value=parseNominal(input.value);
 const schedule=state.schedules.find(x=>x.program_id===program.id&&Number(x.rt)===Number(kk.rt)&&String(x.tanggal).slice(0,4)===String(year)&&Number(String(x.tanggal).slice(5,7))===month+1);
 if(!schedule)return toast('Jadwal bulan ini belum tersedia untuk RT tersebut','err');
 let {data:sessions,error:sessionError}=await supabase.from('penarikan_session').select('*').eq('schedule_id',schedule.id).order('started_at',{ascending:false}).limit(1);
 if(sessionError){toast(sessionError.message,'err');return;}
 let session=sessions?.[0];
 if(!session||session.nama_penarik==='Admin'){const collectorName=state.matrixCollectorNames[`${Number(kk.rt)}:${year}:${month+1}`]||await collectorNameModal();if(!collectorName)return;if(session){const result=await supabase.from('penarikan_session').update({nama_penarik:collectorName,identitas:'Input tabel'}).eq('id',session.id).select().single();if(result.error){toast(result.error.message,'err');return;}session=result.data;}else{const result=await supabase.from('penarikan_session').insert({schedule_id:schedule.id,nama_penarik:collectorName,identitas:'Input tabel',status:'active'}).select().single();if(result.error){toast(result.error.message,'err');return;}session=result.data;}}
 const existing=state.reportTransactions?.find(x=>String(x.kk_id)===String(kk.id)&&String(x.schedule_id)===String(schedule.id));
 const query=existing?supabase.from('jimpitan_transaction').update({nominal:value,updated_at:new Date().toISOString(),updated_by:state.session.user.id}).eq('id',existing.id):supabase.from('jimpitan_transaction').insert({schedule_id:schedule.id,session_id:session.id,kk_id:kk.id,nominal:value,confirmed_by:state.session.user.id,updated_by:state.session.user.id});
 const {error}=await query;if(error){toast(error.message,'err');return;}
 toast('Nominal tersimpan');await loadAll();await loadReport();render();
}
function bindJimpitanMatrix(){
 const table=document.querySelector('.darkMatrix');if(!table)return;
 const headers=[...table.tHead.rows[0].cells],monthHeaders=headers.slice(3,-1),today=new Date(),todayKey=today.getFullYear()*12+today.getMonth(),headerDate=header=>{const text=header.textContent.trim(),month=MONTHS.findIndex(label=>label.slice(0,3).toUpperCase()===text.slice(0,3));return {month,year:Number(text.slice(-4))};};
 const activeIndex=Math.max(0,monthHeaders.findIndex(header=>{const date=headerDate(header);return date.year*12+date.month>=todayKey;}));
 setTimeout(()=>{const wrap=table.closest('.tableWrap'),currentHeader=monthHeaders[activeIndex],stickyWidth=headers.slice(0,3).reduce((total,cell)=>total+cell.getBoundingClientRect().width,0);if(wrap&&currentHeader){const currentLeft=currentHeader.getBoundingClientRect().left-wrap.getBoundingClientRect().left+wrap.scrollLeft;wrap.scrollLeft=Math.max(0,currentLeft-stickyWidth);}},0);
 const collectorCells=[...table.tFoot.querySelector('.collectorRow').cells];
 if(state.session&&state.reportRt!=='all')collectorCells.slice(3,-1).forEach((cell,index)=>{const header=monthHeaders[index],{month,year}=headerDate(header),existing=state.matrixCollectorNames[`${Number(state.reportRt)}:${year}:${month+1}`]||cell.querySelector('.collectorNames')?.textContent.trim()||'';cell.innerHTML=`<div class="collectorNames"><input value="${existing==='—'||existing.toLowerCase()==='admin'?'':esc(existing)}" placeholder="Nama penarik" aria-label="Nama penarik ${header.textContent.trim()}"><button type="button" title="Simpan nama penarik">✓</button></div>`;cell.querySelector('button').onclick=()=>saveMatrixCollectorName(month,year,cell.querySelector('input'));});
 monthHeaders.forEach((header,index)=>{
  const {month,year}=headerDate(header),compact=index<activeIndex;header.classList.toggle('monthCompact',compact);header.title=compact?'Klik untuk membuka bulan':'Bulan aktif';
  header.onclick=()=>{if(index>=activeIndex)return;const next=new Set(state.expandedMonths||[]);if(next.has(index))next.delete(index);else next.add(index);state.expandedMonths=[...next];render();};
  [...table.tBodies[0].rows].forEach(row=>{
   const nameCell=row.cells[1],monthCell=row.cells[index+3];if(!nameCell||!monthCell)return;
   monthCell.classList.toggle('monthCompact',compact);
  if(state.session&&!monthCell.dataset.matrixAmountReady){const text=monthCell.textContent.replace(/[^0-9]/g,'');const value=text?Number(text):0;monthCell.innerHTML=`<div class="matrixAmountEdit"><input type="text" inputmode="numeric" value="${value?formatNominal(value):''}" placeholder="0" aria-label="Nominal bulan ${month}/${year}"><button type="button" title="Simpan nominal">✓</button></div>`;const amountInput=monthCell.querySelector('input');bindNominalInput(amountInput);const name=nameCell.querySelector('input')?.value||nameCell.querySelector('b')?.textContent.trim();const kk=state.kks.find(x=>x.nama===name);if(kk)monthCell.querySelector('button').onclick=()=>saveMatrixNominal(kk.id,year,month,amountInput);monthCell.dataset.matrixAmountReady='true';}
  });
 });
}
function adminKKContent(){
 const selected=state.adminKKRt?state.rts.find(r=>String(r.id)===String(state.adminKKRt)):null;const kks=selected?state.kks.filter(k=>Number(k.rt)===Number(selected.id)):[];
 return `<div class="panelTitle">DATA KK PER RT</div><p class="muted">Pilih RT untuk melihat dan mengelola daftar KK.</p><div class="cardActions"><button class="primary" data-add-kk>+ Tambah KK</button><button class="secondary" id="importKK">Import PDF / Excel / CSV / TXT</button></div><div class="adminRtPills">${state.rts.map(rt=>`<button class="rtPill ${String(state.adminKKRt)===String(rt.id)?'active':''}" data-admin-kk-rt="${rt.id}">${esc(rt.nama)} <small>${state.kks.filter(k=>Number(k.rt)===Number(rt.id)).length}</small></button>`).join('')}</div>${selected?`<div class="adminRtGroup"><div class="adminRtHeader"><b>${esc(selected.nama)}</b><span>${kks.length} KK</span></div><div class="adminList">${kks.map(k=>`<div class="adminItem"><div><b>${esc(k.nama)}</b><span>KK aktif</span></div><div><button class="secondary small" data-edit-kk="${k.id}">Edit</button><button class="dangerText" data-delete-kk="${k.id}">Nonaktifkan</button></div></div>`).join('')||'<div class="muted">Belum ada KK di wilayah ini.</div>'}</div></div>`:'<div class="adminKKEmpty">Klik salah satu pill RT untuk menampilkan data KK.</div>'}`;
}
function editKKModal(id){
 if(!authRequired())return;
 const kk=state.kks.find(x=>String(x.id)===String(id));if(!kk)return;
 modal(`<div class="modalCard compact"><div class="modalHead"><div><div class="eyebrow">DATA KK</div><h3>Edit KK</h3></div><button class="modalClose" data-close-modal>×</button></div><label>Wilayah<select id="editKKRt">${state.rts.map(r=>`<option value="${r.id}" ${Number(r.id)===Number(kk.rt)?'selected':''}>${esc(r.nama)}</option>`).join('')}</select></label><label>Nama KK<input id="editKKName" value="${esc(kk.nama)}" autocomplete="off"></label><div class="modalActions"><button class="secondary" data-close-modal>Batal</button><button class="primary" id="saveEditKK">Simpan</button></div></div>`, 'editKKName');
 document.querySelector('#saveEditKK').onclick=async()=>{const nama=document.querySelector('#editKKName').value.trim(),rt=Number(document.querySelector('#editKKRt').value);if(!nama)return toast('Nama KK wajib diisi','err');const {error}=await supabase.from('kk').update({nama,rt}).eq('id',id);if(error){toast(error.message,'err');return;}closeModal();toast('Data KK diperbarui');await loadKK();render();};
}
function bind(){
 bindJimpitanMatrix();
 document.querySelectorAll('[data-nominal]:not(:disabled)').forEach(bindNominalInput);
 document.querySelector('.jimpitanPage .rtPill[data-report-rt="all"]')?.remove();
 const exportPanel=document.querySelector('.jimpitanPage .jimpitanHelp');
 if(exportPanel){exportPanel.innerHTML='<div class="panelTitle">EXPORT REKAP NOMINAL</div><p>Export per bulan atau rentang bulan. File hanya berisi total nominal tanpa nama KK.</p><button class="secondary" id="openJimpitanExport">EXPORT PDF / CSV</button>';document.querySelector('#openJimpitanExport').onclick=openJimpitanExport;}
 document.querySelectorAll('[data-nav]').forEach(b=>b.onclick=()=>go(b.dataset.nav));
 document.querySelectorAll('[data-start-schedule]').forEach(b=>b.onclick=()=>openSchedule(b.dataset.startSchedule));
 document.querySelectorAll('#newProgram,[data-new-program]').forEach(b=>b.addEventListener('click',()=>programModal()));
 document.querySelectorAll('[data-edit-program]').forEach(b=>b.onclick=()=>programModal(b.dataset.editProgram));
 document.querySelectorAll('[data-toggle-program]').forEach(b=>b.onclick=()=>toggleProgram(b.dataset.toggleProgram));
 document.querySelector('#kkRt')?.addEventListener('change',async e=>{state.rt=e.target.value==='all'?'all':+e.target.value;await loadKK();render();});
 document.querySelector('#kkSearch')?.addEventListener('input',async e=>{state.search=e.target.value;await loadKK();render();});
 document.querySelectorAll('#addKK,[data-add-kk]').forEach(b=>b.addEventListener('click',addKKModal));
 document.querySelector('#importKK')?.addEventListener('click',importKKModal);
 document.querySelectorAll('[data-delete-kk]').forEach(b=>b.onclick=()=>deleteKK(b.dataset.deleteKk));
 document.querySelectorAll('[data-history-kk]').forEach(b=>b.onclick=()=>loadHistory(b.dataset.historyKk));
 document.querySelector('#closeHistory')?.addEventListener('click',()=>{state.kkHistory=null;render();});
 document.querySelectorAll('[data-correct]').forEach(b=>b.onclick=()=>correctTransaction(b.dataset.correct));
 document.querySelector('#scheduleProgram')?.addEventListener('change',e=>{state.selectedProgram=e.target.value;render();});
 document.querySelector('#exportExcel')?.addEventListener('click',exportExcel);
 document.querySelector('#exportPdf')?.addEventListener('click',exportPdf);
 document.querySelector('#applyReport')?.addEventListener('click',async()=>{state.reportYear=+document.querySelector('#reportYear').value;state.reportMonth=document.querySelector('#reportMonth').value;state.reportRt=document.querySelector('#reportRt').value;state.reportProgram=document.querySelector('#reportProgram').value;await loadReport();render();});
 document.querySelectorAll('[data-report-rt]').forEach(b=>b.onclick=async()=>{state.reportRt=b.dataset.reportRt;await loadReport();render();});
 const agendaSearch=document.querySelector('.agendaTools input[placeholder="Cari agenda..."]');
 if(agendaSearch){agendaSearch.disabled=false;agendaSearch.value=state.agendaSearch||'';agendaSearch.oninput=e=>{state.agendaSearch=e.target.value;const query=state.agendaSearch.trim().toLowerCase();document.querySelectorAll('.agendaCard').forEach(card=>{card.hidden=!!query&&!card.textContent.toLowerCase().includes(query);});};}
 document.querySelectorAll('.agendaFilters .filterPill').forEach(button=>button.onclick=()=>{const category=button.textContent.trim().toLowerCase();document.querySelectorAll('.agendaFilters .filterPill').forEach(item=>item.classList.remove('active'));button.classList.add('active');document.querySelectorAll('.agendaCard').forEach(card=>{card.hidden=category!=='semua'&&!card.textContent.toLowerCase().includes(category);});});
 document.querySelector('#reportProgram')?.addEventListener('change',async e=>{state.reportProgram=e.target.value;await loadReport();render();});
 document.querySelector('#reportYear')?.addEventListener('change',async e=>{state.reportYear=+e.target.value;await loadReport();render();});
 document.querySelector('#reportMonth')?.addEventListener('change',async e=>{state.reportMonth=e.target.value;await loadReport();render();});
 document.querySelector('#shareAllAgenda')?.addEventListener('click',()=>{const text=state.agendas.map(a=>`Tanggal: ${a.tgl} ${a.bulan} ${a.tahun}\nJam: ${a.jam}\nTempat: ${a.tempat}\nKegiatan: ${a.acara}\n${a.deskripsi||''}`).join('\n\n');window.open(`https://wa.me/?text=${encodeURIComponent('*Pemuda Pemudi Mawe - Agenda*\n\n'+text)}`,'_blank');});
 document.querySelector('#login')?.addEventListener('click',loginModal);
 document.querySelector('#logout')?.addEventListener('click',async()=>{await supabase.auth.signOut();state.session=null;go('admin');});
 const adminSections={jimpitan:'#adminJimpitan',agenda:'#adminAgenda',kk:'#adminKK',struktur:'#adminStruktur'};const activeAdminTab=state.adminTab||'dashboard';document.querySelectorAll('.adminSection').forEach(section=>section.style.display='none');if(adminSections[activeAdminTab]){const section=document.querySelector(adminSections[activeAdminTab]);if(section){section.style.display='block';if(activeAdminTab==='kk'){const panel=section.querySelector('.panel');if(panel)panel.innerHTML=adminKKContent();}}}document.querySelectorAll('[data-admin-tab]').forEach(button=>{button.classList.toggle('active',button.dataset.adminTab===activeAdminTab);button.onclick=()=>{state.adminTab=button.dataset.adminTab;render();};});
 document.querySelectorAll('[data-edit-kk]').forEach(button=>button.onclick=()=>editKKModal(button.dataset.editKk));
 document.querySelectorAll('[data-admin-kk-rt]').forEach(button=>button.onclick=()=>{state.adminKKRt=button.dataset.adminKkRt;render();});
 document.querySelectorAll('#adminKK [data-add-kk]').forEach(button=>button.onclick=addKKModal);document.querySelectorAll('#adminKK [data-delete-kk]').forEach(button=>button.onclick=()=>deleteKK(button.dataset.deleteKk));document.querySelector('#adminKK #importKK')?.addEventListener('click',importKKModal);
 document.querySelector('#saveOrg')?.addEventListener('click',saveOrg);
 document.querySelectorAll('#newAgenda,[data-new-agenda]').forEach(b=>b.addEventListener('click',()=>agendaModal()));
 document.querySelectorAll('[data-del-agenda]').forEach(b=>b.onclick=()=>deleteAgenda(b.dataset.delAgenda));
 document.querySelectorAll('[data-edit-agenda]').forEach(b=>b.onclick=()=>agendaModal(b.dataset.editAgenda));
 document.querySelectorAll('[data-wa-agenda]').forEach(b=>b.onclick=()=>shareAgenda(b.dataset.waAgenda));
 document.querySelectorAll('[data-confirm-kk]').forEach(b=>b.onclick=()=>confirmTransaction(b.dataset.confirmKk));
 document.querySelector('#backPenarikan')?.addEventListener('click',()=>go('penarikan'));
 document.querySelector('#finishCollection')?.addEventListener('click',finishCollection);
}
async function programModal(id){
 if(!authRequired())return;
 const p=id?state.programs.find(x=>String(x.id)===String(id)):null;
 modal(`
  <div class="modalCard wideModal">
   <div class="modalHead"><div><div class="eyebrow">JIMPITAN</div><h3>${id?'Edit Program':'Tambah Program Jimpitan'}</h3><p class="muted">Admin menentukan program dan jadwal. Penarik ditentukan saat mulai penarikan.</p></div><button class="modalClose" data-close-modal>×</button></div>
   <div class="twoCol">
    <label>Nama program<input id="jpNama" value="${esc(p?.nama_program||'')}" placeholder="Jimpitan Kebersihan"></label>
    <label>Jenis jimpitan<input id="jpJenis" value="${esc(p?.jenis||'')}" placeholder="Kebersihan"></label>
    <label>Target nominal per KK<input id="jpTarget" type="number" min="0" value="${p?.target_nominal??200000}"></label>
    <label>Frekuensi<select id="jpFreq"><option value="weekly">Setiap minggu</option><option value="biweekly">Setiap 2 minggu</option><option value="monthly">Setiap bulan</option><option value="custom">Custom</option></select></label>
    <label>Tanggal mulai<input id="jpMulai" type="date" value="${p?.tanggal_mulai||'2027-01-05'}"></label>
    <label>Tanggal selesai<input id="jpSelesai" type="date" value="${p?.tanggal_selesai||'2027-10-31'}"></label>
   </div>
   <label id="customDatesWrap" class="hiddenField">Tanggal custom <input id="jpCustom" placeholder="2027-01-05, 2027-01-20"></label>
   <div class="modalSectionTitle">Wilayah yang ikut</div>
   <div class="checkGrid">${state.rts.map(r=>`<label class="checkPill"><input type="checkbox" class="jpRt" value="${r.id}" ${(p?.rt_ids||state.rts.map(x=>x.id)).map(Number).includes(Number(r.id))?'checked':''}><span>${esc(r.nama)}</span></label>`).join('')}</div>
   <div class="modalActions"><button class="secondary" data-close-modal>Batal</button><button class="primary" id="saveProgramModal">${id?'Simpan Perubahan':'Buat Program & Jadwal'}</button></div>
  </div>`, 'jpNama');
 const freq=document.querySelector('#jpFreq');freq.value=p?.frekuensi||'biweekly';const toggle=()=>document.querySelector('#customDatesWrap').classList.toggle('show',freq.value==='custom');freq.onchange=toggle;toggle();
 document.querySelector('#saveProgramModal').onclick=async()=>{
  const nama=document.querySelector('#jpNama').value.trim(),jenis=document.querySelector('#jpJenis').value.trim(),target=Number(document.querySelector('#jpTarget').value),mulai=document.querySelector('#jpMulai').value,selesai=document.querySelector('#jpSelesai').value,freqValue=freq.value;
  const rt_ids=[...document.querySelectorAll('.jpRt:checked')].map(x=>Number(x.value));
  if(!nama||!jenis||!mulai||!selesai||!rt_ids.length)return toast('Lengkapi program, jenis, tanggal dan minimal satu wilayah.','err');
  if(!Number.isFinite(target)||target<0)return toast('Target nominal tidak valid','err');
  if(selesai<mulai)return toast('Tanggal selesai harus setelah tanggal mulai','err');
  const custom_dates=freqValue==='custom'?document.querySelector('#jpCustom').value.split(',').map(x=>x.trim()).filter(Boolean):[];
  if(freqValue==='custom'&&!custom_dates.length)return toast('Isi minimal satu tanggal custom','err');
  const payload={nama_program:nama,jenis,target_nominal:target,tanggal_mulai:mulai,tanggal_selesai:selesai,frekuensi:freqValue,custom_dates,rt_ids,aktif:p?.aktif??true,created_by:p?.created_by||state.session.user.id};
  const q=p?supabase.from('jimpitan_program').update(payload).eq('id',p.id):supabase.from('jimpitan_program').insert(payload).select().single();
  const {data,error}=await q;if(error)return toast(error.message,'err');
  const program=data||p;await syncSchedules(program.id,payload);
  toast(`Program tersimpan. ${generateDates(payload).length*rt_ids.length} jadwal disiapkan.`);closeModal();await loadAll();
 };
}
async function syncSchedules(programId,p){
 const {data:existing,error}=await supabase.from('jimpitan_schedule').select('id,tanggal,rt,status').eq('program_id',programId);
 if(error){toast(error.message,'err');return;}
 const dates=generateDates(p),wanted=new Set();
 for(const d of dates)for(const rt of (p.rt_ids||[]))wanted.add(`${d}|${rt}`);
 const keep=(existing||[]).filter(s=>s.status==='finished'||s.status==='active').map(s=>`${s.tanggal}|${s.rt}`);
 const deletable=(existing||[]).filter(s=>s.status==='scheduled'&&!wanted.has(`${s.tanggal}|${s.rt}`)).map(s=>s.id);
 if(deletable.length)await supabase.from('jimpitan_schedule').delete().in('id',deletable);
 const current=new Set((existing||[]).map(s=>`${s.tanggal}|${s.rt}`));
 const rows=[];for(const key of wanted){if(!current.has(key)&&!keep.includes(key)){const [tanggal,rt]=key.split('|');rows.push({program_id:programId,tanggal,rt:Number(rt),status:'scheduled'});}}
 if(rows.length){const {error:e}=await supabase.from('jimpitan_schedule').insert(rows);if(e)toast(e.message,'err');}
}

function generateDates(p){
 const out=[];
 if(p.frekuensi==='custom')return [...new Set((p.custom_dates||[]).filter(d=>d>=p.tanggal_mulai&&d<=p.tanggal_selesai))].sort();
 let d=new Date(`${p.tanggal_mulai}T00:00:00`),end=new Date(`${p.tanggal_selesai}T00:00:00`);
 while(d<=end){out.push(d.toISOString().slice(0,10));if(p.frekuensi==='weekly')d.setDate(d.getDate()+7);else if(p.frekuensi==='biweekly')d.setDate(d.getDate()+14);else d.setMonth(d.getMonth()+1);}
 return out;
}
async function toggleProgram(id){
 if(!authRequired())return;const p=state.programs.find(x=>String(x.id)===String(id));if(!p)return;
 const {error}=await supabase.from('jimpitan_program').update({aktif:!p.aktif}).eq('id',id);if(error)toast(error.message,'err');else{toast(p.aktif?'Program dinonaktifkan':'Program diaktifkan');await loadAll();}
}
function addKKModal(){
 if(!authRequired())return;
 modal(`
  <div class="modalCard compact">
   <div class="modalHead"><div><div class="eyebrow">DATA KK</div><h3>Tambah KK</h3></div><button class="modalClose" data-close-modal>×</button></div>
   <label>Wilayah<select id="modalKKRt">${state.rts.map(r=>`<option value="${r.id}" ${Number(state.rt)===Number(r.id)?'selected':''}>${esc(r.nama)}</option>`).join('')}</select></label>
   <label>Nama KK<input id="modalKKName" autocomplete="off" placeholder="Contoh: Ahmad" autofocus></label>
   <div class="modalActions"><button class="secondary" data-close-modal>Batal</button><button class="primary" id="saveKKModal">Simpan KK</button></div>
  </div>`, 'saveKKModal');
 document.querySelector('#saveKKModal').onclick=async()=>{const name=document.querySelector('#modalKKName').value.trim();const rt=Number(document.querySelector('#modalKKRt').value);if(!name)return toast('Nama KK wajib diisi','err');await insertKK(name,rt);closeModal();};
 document.querySelector('#modalKKName').onkeydown=e=>{if(e.key==='Enter')document.querySelector('#saveKKModal').click();};
}
async function insertKK(name,rt){if(!authRequired())return;const {error}=await supabase.from('kk').insert({nama:name.trim(),rt,aktif:true});if(error)toast(error.message,'err');else{toast('KK ditambahkan');await loadKK();render();}}
async function deleteKK(id){if(!authRequired())return;if(!confirm('Nonaktifkan KK ini? Riwayat transaksi tetap disimpan.'))return;const {error}=await supabase.from('kk').update({aktif:false}).eq('id',id);if(error)toast(error.message,'err');else{toast('KK dinonaktifkan; riwayat tetap aman.');await loadKK();render();}}
async function loadHistory(id){
 const {data:kk}=await supabase.from('kk').select('*').eq('id',id).single();if(!kk)return;
 const {data,error}=await supabase.from('jimpitan_transaction').select('*, kk(nama,rt), penarikan_session(nama_penarik), jimpitan_schedule(tanggal,jimpitan_program(nama_program,jenis))').eq('kk_id',id).order('confirmed_at',{ascending:false});
 if(error){toast(error.message,'err');return;}state.kkHistory={...kk,rows:data||[]};go('penarikan');
}
async function correctTransaction(id){
 if(!authRequired())return;
 const {data:t}=await supabase.from('jimpitan_transaction').select('*').eq('id',id).single();if(!t)return;
 modal(`
  <div class="modalCard compact">
   <div class="modalHead"><div><div class="eyebrow">KOREKSI TRANSAKSI</div><h3>Perbaiki Nominal</h3><p class="muted">Nominal awal: ${money(t.nominal)}</p></div><button class="modalClose" data-close-modal>×</button></div>
   <label>Nominal baru<input id="correctionNominal" type="number" min="0" step="1000" value="${Number(t.nominal||0)}"></label>
   <label>Alasan koreksi<textarea id="correctionReason" rows="3">Koreksi nominal</textarea></label>
   <div class="modalActions"><button class="secondary" data-close-modal>Batal</button><button class="primary" id="saveCorrection">Simpan Koreksi</button></div>
  </div>`, 'correctionNominal');
 document.querySelector('#saveCorrection').onclick=async()=>{
  const next=Number(document.querySelector('#correctionNominal').value);
  const reason=document.querySelector('#correctionReason').value.trim();
  if(!Number.isFinite(next)||next<0)return toast('Nominal tidak valid','err');
  if(!reason)return toast('Alasan koreksi wajib diisi','err');
  const {error:e}=await supabase.from('jimpitan_correction').insert({transaction_id:id,old_nominal:t.nominal,new_nominal:next,reason,corrected_by:state.session.user.id});
  if(e)return toast(e.message,'err');
  const {error}=await supabase.from('jimpitan_transaction').update({nominal:next,updated_at:new Date().toISOString(),updated_by:state.session.user.id}).eq('id',id);
  if(error){toast(error.message,'err');return;}
  closeModal();toast('Transaksi dikoreksi dan riwayat perubahan tersimpan.');
  if(state.kkHistory)await loadHistory(id);
 };
}
async function confirmTransaction(kkId){
 const s=state.schedules.find(x=>String(x.id)===String(state.selectedSchedule));const session=state.sessions[0];if(!s||!session)return;
 const input=document.querySelector(`[data-nominal="${kkId}"]`);const nominal=parseNominal(input?.value);
 const {data:existing}=await supabase.from('jimpitan_transaction').select('id').eq('schedule_id',s.id).eq('kk_id',kkId).maybeSingle();
 if(existing)return toast('KK ini sudah dicatat.','err');
 const {error}=await supabase.from('jimpitan_transaction').insert({schedule_id:s.id,session_id:session.id,kk_id:kkId,nominal,confirmed_at:new Date().toISOString(),confirmed_by:null});
 if(error)toast(error.message,'err');else{await loadScheduleTransactions(s.id);render();}
}
async function finishCollection(){
 const s=state.schedules.find(x=>String(x.id)===String(state.selectedSchedule));const session=state.sessions[0];if(!s||!session)return;
 if(!confirm('Selesaikan penarikan ini? KK yang belum dikonfirmasi tetap belum memiliki transaksi pada jadwal ini.'))return;
 const {error}=await supabase.from('penarikan_session').update({finished_at:new Date().toISOString(),status:'finished'}).eq('id',session.id);
 if(error)return toast(error.message,'err');
 await supabase.from('jimpitan_schedule').update({status:'finished'}).eq('id',s.id);state.sessions=[];state.selectedSchedule='';await loadAll();go('penarikan');
}
async function loadReport(dateRange=null){
 if(!supabase)return;
 let q=supabase.from('jimpitan_transaction').select('*, kk(nama,rt), penarikan_session(nama_penarik), jimpitan_schedule(tanggal,program_id,jimpitan_program(nama_program,jenis,target_nominal))');
 const matrixProgram=state.view==='jimpitan'&&state.reportProgram!=='all'?state.programs.find(x=>String(x.id)===String(state.reportProgram)):null;
 if(dateRange){const endDate=new Date(dateRange.endYear,dateRange.endMonth,0).toISOString().slice(0,10);q=q.gte('jimpitan_schedule.tanggal',`${dateRange.startYear}-${String(dateRange.startMonth).padStart(2,'0')}-01`).lte('jimpitan_schedule.tanggal',endDate);}
 else if(matrixProgram?.tanggal_mulai&&matrixProgram?.tanggal_selesai)q=q.gte('jimpitan_schedule.tanggal',matrixProgram.tanggal_mulai).lte('jimpitan_schedule.tanggal',matrixProgram.tanggal_selesai);
 else if(state.reportYear)q=q.gte('jimpitan_schedule.tanggal',`${state.reportYear}-01-01`).lte('jimpitan_schedule.tanggal',`${state.reportYear}-12-31`);
 const {data,error}=await q.order('confirmed_at',{ascending:true});if(error){toast(error.message,'err');return;}
 state.reportTransactions=(data||[]).filter(x=>{
   const s=x.jimpitan_schedule||{};const month=s.tanggal?Number(s.tanggal.slice(5,7)):0;
   return (state.reportMonth==='all'||Number(state.reportMonth)===month)&&(state.reportRt==='all'||Number(x.kk?.rt)===Number(state.reportRt))&&(state.reportProgram==='all'||String(s.program_id)===String(state.reportProgram));
 });
}
async function exportExcel(){
 if(!state.reportTransactions?.length)await loadReport();
 const rows=[['Tanggal','RT','KK','Jenis','Program','Penarik','Nominal','Jam']];
 (state.reportTransactions||[]).forEach(x=>rows.push([x.jimpitan_schedule?.tanggal,rtName(x.kk?.rt),x.kk?.nama,x.jimpitan_schedule?.jimpitan_program?.jenis,x.jimpitan_schedule?.jimpitan_program?.nama_program,x.penarikan_session?.nama_penarik,Number(x.nominal||0),timeNow(x.confirmed_at)]));
 const ws=XLSX.utils.aoa_to_sheet(rows),wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Transaksi');XLSX.writeFile(wb,`laporan-jimpitan-${state.reportYear}.xlsx`);
}
function exportPdf(){window.print();}
function jimpitanExportRows(startMonth,startYear,endMonth,endYear,rtFilter=state.reportRt){
 const transactions=state.reportTransactions||[],selectedRt=rtFilter==='all'?'all':Number(rtFilter),households=state.kks.filter(kk=>selectedRt==='all'||Number(kk.rt)===selectedRt),rows=[];
 const householdIds=new Set(households.map(kk=>String(kk.id)));
 transactions.forEach(transaction=>{if(!householdIds.has(String(transaction.kk_id))){const kk=transaction.kk;if(kk&&(selectedRt==='all'||Number(kk.rt)===selectedRt))households.push(kk);}});
 for(let cursor=startYear*12+startMonth-1;cursor<=endYear*12+endMonth-1;cursor++){const year=Math.floor(cursor/12),month=cursor%12+1;households.forEach(kk=>{
  const nominal=transactions.filter(transaction=>String(transaction.kk_id)===String(kk.id)&&Number(String(transaction.jimpitan_schedule?.tanggal||'').slice(0,4))===year&&Number(String(transaction.jimpitan_schedule?.tanggal||'').slice(5,7))===month).reduce((sum,transaction)=>sum+Number(transaction.nominal||0),0);
  rows.push({month,year,monthName:`${MONTHS[month-1]} ${year}`,rt:rtName(kk.rt),kkName:kk.nama||'—',nominal});
 });}
 return rows.sort((a,b)=>a.year-b.year||a.month-b.month||a.rt.localeCompare(b.rt)||a.kkName.localeCompare(b.kkName));
}
async function openJimpitanExport(){
 if(!state.reportTransactions?.length)await loadReport();
 modal(`<div class="modalCard compact"><div class="modalHead"><div><div class="eyebrow">EXPORT JIMPITAN</div><h3>Rekap Nominal</h3><p class="muted">Dengan nama KK. Pilih RT dan rentang bulan.</p></div><button class="modalClose" data-close-modal>×</button></div><label>RT yang diexport<select id="exportRt"><option value="all">Semua RT</option>${state.rts.map(r=>`<option value="${r.id}" ${String(state.reportRt)===String(r.id)?'selected':''}>${esc(r.nama)}</option>`).join('')}</select></label><label class="exportAllLabel"><input id="exportAllMonths" type="checkbox"><span>Semua bulan tahun ${state.reportYear}</span></label><div class="twoCol"><label>Dari bulan<select id="exportStartMonth">${MONTHS.map((m,i)=>`<option value="${i+1}">${m}</option>`).join('')}</select></label><label>Tahun<select id="exportStartYear">${[state.reportYear,state.reportYear+1].map(y=>`<option value="${y}">${y}</option>`).join('')}</select></label><label>Sampai bulan<select id="exportEndMonth">${MONTHS.map((m,i)=>`<option value="${i+1}" ${i===11?'selected':''}>${m}</option>`).join('')}</select></label><label>Tahun<select id="exportEndYear"><option value="${state.reportYear}">${state.reportYear}</option><option value="${state.reportYear+1}" selected>${state.reportYear+1}</option></select></label></div><div class="modalActions"><button class="secondary" data-close-modal>Batal</button><button class="secondary" id="exportJimpitanCsv">CSV</button><button class="primary" id="exportJimpitanPdf">PDF A4</button></div></div>`, 'exportRt');
 const getRange=()=>{if(document.querySelector('#exportAllMonths').checked)return [1,state.reportYear,12,state.reportYear];const start=Number(document.querySelector('#exportStartMonth').value),startYear=Number(document.querySelector('#exportStartYear').value),end=Number(document.querySelector('#exportEndMonth').value),endYear=Number(document.querySelector('#exportEndYear').value);if(startYear*12+start>endYear*12+end){toast('Bulan mulai tidak boleh setelah bulan akhir','err');return null;}return [start,startYear,end,endYear];};
 document.querySelector('#exportAllMonths').onchange=e=>{document.querySelector('#exportStartMonth').disabled=e.target.checked;document.querySelector('#exportEndMonth').disabled=e.target.checked;};
 const exportWithRt=async(type)=>{const range=getRange();if(!range)return;const selectedRt=document.querySelector('#exportRt').value,previousRt=state.reportRt,previousMonth=state.reportMonth;state.reportRt=selectedRt;state.reportMonth='all';const [startMonth,startYear,endMonth,endYear]=range;await loadReport({startMonth,startYear,endMonth,endYear});if(type==='csv')exportJimpitanCsv(...range,selectedRt);else exportJimpitanPdf(...range,selectedRt);state.reportRt=previousRt;state.reportMonth=previousMonth;await loadReport();};
 document.querySelector('#exportJimpitanCsv').onclick=()=>exportWithRt('csv');
 document.querySelector('#exportJimpitanPdf').onclick=()=>exportWithRt('pdf');
}
function exportJimpitanCsv(startMonth,startYear,endMonth,endYear,rtFilter=state.reportRt){
 const rows=jimpitanExportRows(startMonth,startYear,endMonth,endYear,rtFilter),rt=rtFilter==='all'?'Semua RT':rtName(rtFilter),csv=[['Rekap Jimpitan',`${rt} - ${startYear} sampai ${endYear}`],['Bulan','RT','Nama KK','Nominal'],...rows.map(row=>[row.monthName,row.rt,row.kkName,row.nominal])].map(row=>row.map(value=>`"${String(value).replace(/"/g,'""')}"`).join(',')).join('\r\n');
 const link=document.createElement('a');link.href=URL.createObjectURL(new Blob([`\uFEFF${csv}`],{type:'text/csv;charset=utf-8'}));link.download=`rekap-jimpitan-${startYear}-${startMonth}-${endYear}-${endMonth}.csv`;link.click();URL.revokeObjectURL(link.href);closeModal();
}
function exportJimpitanPdf(startMonth,startYear,endMonth,endYear,rtFilter=state.reportRt){
 const rows=jimpitanExportRows(startMonth,startYear,endMonth,endYear,rtFilter),rt=rtFilter==='all'?'Semua RT':rtName(rtFilter),program=state.reportProgram==='all'?'Semua Program':state.programs.find(x=>String(x.id)===String(state.reportProgram))?.nama_program||'Program';
 const months=[];for(let cursor=startYear*12+startMonth-1;cursor<=endYear*12+endMonth-1;cursor++){const year=Math.floor(cursor/12),month=cursor%12+1;months.push({year,month,label:`${MONTHS[month-1]} ${year}`,key:`${year}-${month}`});}
 const groups=new Map();rows.forEach(row=>{const key=`${row.rt}|${row.kkName}`,group=groups.get(key)||{rt:row.rt,kkName:row.kkName,values:{},total:0};group.values[`${row.year}-${row.month}`]=row.nominal;group.total+=row.nominal;groups.set(key,group);});
 const groupRows=[...groups.values()],doc=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'}),formatMoney=value=>`Rp ${new Intl.NumberFormat('id-ID').format(value)}`,page={left:7,right:290,top:27,bottom:197},fixedWidth=8+18+48+20,monthWidth=(page.right-page.left-fixedWidth)/months.length,headerHeight=8,rowHeight=Math.max(1.4,Math.min(6,(page.bottom-page.top-headerHeight-8)/(groupRows.length+1))),fontSize=rowHeight<2.5?3:5,total=groupRows.reduce((sum,row)=>sum+row.total,0);
 doc.setFont('helvetica','bold');doc.setFontSize(15);doc.text('Rekap Jimpitan',page.left,10);doc.setFont('helvetica','normal');doc.setFontSize(7);doc.text(`${rt} | ${program} | ${startYear} sampai ${endYear}`,page.left,16);doc.setFontSize(6);doc.text('Nominal transaksi asli per bulan',page.left,21);
 const columns=[{label:'No',width:8},{label:'RT',width:18},{label:'Nama KK',width:48},...months.map(month=>({label:MONTHS[month.month-1],year:String(month.year),width:monthWidth})),{label:'TOTAL',width:20}];let x=page.left;doc.setTextColor(0,0,0);doc.setFont('helvetica','bold');doc.setFontSize(Math.min(fontSize,4.5));columns.forEach(column=>{doc.setFillColor(255,255,255);doc.setDrawColor(70,70,70);doc.rect(x,page.top,column.width,headerHeight,'FD');const headerText=column.year?`${column.label} ${column.year}`:column.label;doc.text(headerText,x+column.width/2,page.top+headerHeight*.68,{align:'center',maxWidth:column.width-1});x+=column.width;});doc.setTextColor(0,0,0);
 groupRows.forEach((row,index)=>{let y=page.top+headerHeight+rowHeight*index;x=page.left;const cells=[String(index+1),row.rt,row.kkName,...months.map(month=>row.values[month.key]?formatMoney(row.values[month.key]):'—'),formatMoney(row.total)];columns.forEach((column,columnIndex)=>{doc.setFont('helvetica',columnIndex===columns.length-1?'bold':'normal');doc.rect(x,y,column.width,rowHeight);doc.text(cells[columnIndex],x+(columnIndex===2?1:column.width/2),y+rowHeight*.68,{align:columnIndex===2?'left':'center',maxWidth:column.width-2});x+=column.width;});});
 const totalY=Math.min(page.bottom+5,page.top+headerHeight+rowHeight*groupRows.length+5);doc.setFont('helvetica','bold');doc.setFontSize(7);doc.text(`TOTAL PERIODE: ${formatMoney(total)} | ${groupRows.length} KK`,page.left,totalY);doc.save(`rekap-jimpitan-${startYear}-${startMonth}-${endYear}-${endMonth}.pdf`);closeModal();
}
function importKKModal(){
 if(!authRequired())return;
 const input=document.createElement('input');
 input.type='file';
 input.accept='.txt,.csv,.xlsx,.xls,.pdf';
 input.onchange=()=>{const file=input.files?.[0];if(file)selectImportRTModal(file);};
 input.click();
}
function selectImportRTModal(file){
 modal(`
  <div class="modalCard compact">
   <div class="modalHead"><div><div class="eyebrow">IMPORT DATA KK</div><h3>Pilih Wilayah Tujuan</h3><p class="muted">${esc(file.name)}</p></div><button class="modalClose" data-close-modal>×</button></div>
   <label>RT / Wilayah<select id="importKKRt">${state.rts.map(r=>`<option value="${r.id}" ${state.rt!=='all'&&Number(state.rt)===Number(r.id)?'selected':''}>${esc(r.nama)}</option>`).join('')}</select></label>
   <div class="modalActions"><button class="secondary" data-close-modal>Batal</button><button class="primary" id="confirmImportKK">Import Data</button></div>
  </div>`);
 document.querySelector('#confirmImportKK').onclick=async()=>{
  const rt=Number(document.querySelector('#importKKRt').value);
  closeModal();
  await parseImport(file,rt);
 };
}
async function parseImport(file,rt){
 if(!file||!rt)return;
 try{
  let names=[];
  const lower=file.name.toLowerCase();
  if(lower.endsWith('.txt')||lower.endsWith('.csv')){
   const text=await file.text();
   names=parseTextNames(text);
  } else if(/\.(xlsx|xls)$/i.test(file.name)){
   const wb=XLSX.read(await file.arrayBuffer());
   const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,raw:false});
   names=rows.flat().map(x=>String(x??'').trim()).filter(x=>x&&!/^nama( kk)?$/i.test(x));
  } else if(lower.endsWith('.pdf')){
   const buf=await file.arrayBuffer();
   const pdf=await pdfjsLib.getDocument({data:buf}).promise;
   names=[];
   for(let i=1;i<=pdf.numPages;i++){
    const page=await pdf.getPage(i);
    const content=await page.getTextContent();
    names.push(...parsePdfPageNames(content.items));
   }
   toast(`PDF terbaca: ${names.length} baris nama.`, 'ok');
  } else {
    return toast('Format file tidak didukung. Gunakan TXT, PDF, Excel, atau CSV.','err');
  }

  // Hapus nomor/baris header, tetapi JANGAN melakukan Set() global.
  // Dua KK berbeda boleh memiliki nama yang sama dan tetap harus menjadi 2 record.
  names=names
   .map(x=>String(x).replace(/^\s*\d+[.)-]\s*/,'').trim())
   .filter(Boolean)
   .filter(x=>!isKKHeader(x));

  if(!names.length)return toast('Tidak menemukan nama KK. Periksa format TXT/PDF/Excel/CSV.','err');

  const {error}=await supabase.from('kk').insert(names.map(n=>({nama:n,rt,aktif:true})));
  if(error)throw error;
  toast(`${names.length} KK berhasil diimport ke ${rtName(rt)}.`);
  await loadKK();
  render();
 }catch(e){
  console.error('Import KK gagal:',e);
  toast(e.message||'Import gagal','err');
 }
}

function isKKHeader(value){
 const x=String(value||'').trim().toLowerCase();
 return !x || /^(daftar\s+nama|nama(\s+kk)?|no(\.|mor)?|nomor|kk|alamat|rt|rw)$/.test(x);
}

function parseTextNames(text){
 return String(text||'')
  .split(/\r?\n/)
  .map(x=>x.replace(/[;,|].*$/,'').trim())
  .filter(x=>x.length>1&&!isKKHeader(x));
}

// PDF.js mengembalikan text item yang kadang terpisah-pisah.
// Jangan menggabungkan seluruh halaman dengan join(' '), karena itu membuat
// semua KK dalam satu halaman menjadi satu record. Kelompokkan item berdasarkan
// posisi Y sehingga 1 baris visual PDF = 1 KK, lalu gabungkan potongan kata
// yang memang berada pada baris yang sama.
function parsePdfPageNames(items){
 const rows=[];
 const Y_TOLERANCE=3;

 for(const item of items||[]){
  const value=String(item?.str||'').replace(/\s+/g,' ').trim();
  if(!value)continue;
  const transform=item.transform||[];
  const x=Number(transform[4]||0);
  const y=Number(transform[5]||0);

  let row=rows.find(r=>Math.abs(r.y-y)<=Y_TOLERANCE);
  if(!row){
   row={y,items:[]};
   rows.push(row);
  }
  row.items.push({x,value});
 }

 return rows
  .sort((a,b)=>b.y-a.y)
  .map(row=>row.items.sort((a,b)=>a.x-b.x).map(x=>x.value).join(' ').replace(/\s+/g,' ').trim())
  .filter(x=>x.length>1&&!isKKHeader(x));
}
async function saveOrg(){
 if(!authRequired())return;const o=state.org||defaultOrg(),koordinator={},bidang={};state.rts.forEach(r=>koordinator[r.id]=document.querySelector(`#org_rt_${r.id}`)?.value.trim()||'');['keamanan','olahraga','seni_budaya','sosial','humas','kerohanian','lingkungan'].forEach(k=>bidang[k]=document.querySelector(`#org_bidang_${k}`)?.value.trim()||'');
 const payload={id:1,ketua:document.querySelector('#org_ketua').value.trim(),wakil:document.querySelector('#org_wakil').value.trim(),sekretaris:document.querySelector('#org_sekretaris').value.trim(),bendahara:document.querySelector('#org_bendahara').value.trim(),koordinator,visi:o.visi,misi:o.misi,updated_at:new Date().toISOString()};
 const extended={...payload,pembina:document.querySelector('#org_pembina')?.value.trim()||'',sekretaris2:document.querySelector('#org_sekretaris2')?.value.trim()||'',bendahara2:document.querySelector('#org_bendahara2')?.value.trim()||'',bidang};
 const {error}=await supabase.from('organisasi').upsert(extended);if(error){toast(error.message,'err');return;}state.org=extended;toast('Struktur organisasi tersimpan');render();
}
function closeModal(){document.querySelector('#appModal')?.remove();}
function modal(inner,focusId){
 closeModal();
 const el=document.createElement('div');el.id='appModal';el.className='appModal';el.innerHTML=`<div class="modalBackdrop" data-close-modal></div><div class="modalWrap">${inner}</div>`;document.body.appendChild(el);
 el.querySelectorAll('[data-close-modal]').forEach(x=>x.addEventListener('click',e=>{if(e.target===x||x.dataset.closeModal!==undefined)closeModal();}));
 el.querySelectorAll('input[type="date"],input[type="time"]').forEach(input=>input.addEventListener('click',()=>{if(typeof input.showPicker==='function'){try{input.showPicker();}catch{}}}));
 if(focusId)setTimeout(()=>document.getElementById(focusId)?.focus(),50);
}

function agendaModal(id){
 if(!authRequired())return;
 const a=id?state.agendas.find(x=>String(x.id)===String(id)):{};
 const dateValue=a?.tahun&&a?.tgl?`${a.tahun}-${String(MONTHS.indexOf(a.bulan)+1).padStart(2,'0')}-${String(a.tgl).padStart(2,'0')}`:'';
 modal(`
  <div class="modalCard">
   <div class="modalHead"><div><div class="eyebrow">AGENDA</div><h3>${id?'Edit Agenda':'Buat Agenda'}</h3></div><button class="modalClose" data-close-modal>×</button></div>
   <div class="twoCol">
    <label>Tanggal<input id="agendaDate" type="date" value="${dateValue||new Date().toISOString().slice(0,10)}"></label>
    <label>Jam<input id="agendaJam" type="time" value="${(a?.jam||'19:00').slice(0,5)}"></label>
   </div>
   <label>Nama kegiatan<input id="agendaAcara" value="${esc(a?.acara||'')}" placeholder="Contoh: Rapat Pemuda"></label>
   <label>Tempat<input id="agendaTempat" value="${esc(a?.tempat||'')}" placeholder="Contoh: Balai Desa"></label>
   <label>Deskripsi<textarea id="agendaDeskripsi" rows="4" placeholder="Keterangan kegiatan...">${esc(a?.deskripsi||'')}</textarea></label>
   <div class="modalActions"><button class="secondary" data-close-modal>Batal</button><button class="primary" id="saveAgendaModal">${id?'Simpan Perubahan':'Buat Agenda'}</button></div>
  </div>`, 'saveAgendaModal');
 document.querySelector('#saveAgendaModal').onclick=async()=>{
  const date=document.querySelector('#agendaDate').value, acara=document.querySelector('#agendaAcara').value.trim(), tempat=document.querySelector('#agendaTempat').value.trim();
  if(!date||!acara||!tempat)return toast('Tanggal, nama kegiatan dan tempat wajib diisi','err');
  const d=new Date(`${date}T00:00:00`), jam=document.querySelector('#agendaJam').value||'19:00';
  await saveAgenda({id:a?.id,tgl:d.getDate(),bulan:MONTHS[d.getMonth()],tahun:d.getFullYear(),acara,tempat,jam:jam+' WIB',deskripsi:document.querySelector('#agendaDeskripsi').value.trim()});
 };
}
async function saveAgenda(payload){const q=payload.id?supabase.from('agenda').update(payload).eq('id',payload.id):supabase.from('agenda').insert(payload);const {error}=await q;if(error)toast(error.message,'err');else{toast('Agenda tersimpan');await loadAll();closeModal();go('agenda');}}
async function deleteAgenda(id){if(!authRequired())return;if(!confirm('Hapus agenda ini?'))return;const {error}=await supabase.from('agenda').delete().eq('id',id);if(error)toast(error.message,'err');else{toast('Agenda dihapus');await loadAll();}}
function shareAgenda(id){const a=state.agendas.find(x=>String(x.id)===String(id));if(!a)return;const txt=`*Pemuda Pemudi Mawe - Agenda*\n\nTanggal: ${a.tgl} ${a.bulan} ${a.tahun}\nJam: ${a.jam}\nTempat: ${a.tempat}\nKegiatan: ${a.acara}\n\n${a.deskripsi||''}\n\nGuyub, Rukun, Sejahtera!`;window.open(`https://wa.me/?text=${encodeURIComponent(txt)}`,'_blank');}
function loginModal(){
 if(!requireConfig())return;
 document.querySelector('#loginOverlay')?.remove();
 const el=document.createElement('div'); el.id='loginOverlay'; el.className='loginOverlay';
 el.innerHTML=`<div class="loginCard"><button class="loginClose" id="loginClose">×</button><div class="eyebrow">ADMIN ACCESS</div><h3>Login Admin</h3><p class="muted">Akun admin sudah terisi otomatis. Masukkan password saja.</p><label>Email Admin<input id="loginEmail" type="email" autocomplete="username" value="${esc(ADMIN_EMAIL)}" readonly placeholder="admin@pemudamawe.com"></label><label>Password<input id="loginPassword" type="password" autocomplete="current-password" placeholder="••••••••"></label><button class="primary loginSubmit" id="loginSubmit">Masuk sebagai Admin</button><div id="loginError" class="loginError"></div></div>`;
 document.body.appendChild(el);
 const close=()=>el.remove(); document.querySelector('#loginClose').onclick=close; el.addEventListener('click',e=>{if(e.target===el)close();});
 const submit=async()=>{const email=document.querySelector('#loginEmail').value.trim() || ADMIN_EMAIL; const password=document.querySelector('#loginPassword').value; if(!email||!password){document.querySelector('#loginError').textContent='Password wajib diisi.';return;} const btn=document.querySelector('#loginSubmit');btn.disabled=true;btn.textContent='Memeriksa…';document.querySelector('#loginError').textContent=''; const ok=await doLogin(email,password);if(ok)close();else{btn.disabled=false;btn.textContent='Masuk sebagai Admin';}};
 document.querySelector('#loginSubmit').onclick=submit; document.querySelector('#loginPassword').onkeydown=e=>{if(e.key==='Enter')submit();}; setTimeout(()=>document.querySelector('#loginPassword')?.focus(),50);
}
async function doLogin(email,password){
 if(!requireConfig())return false;
 const raw=(email||ADMIN_EMAIL).trim();
 const normalized = raw.includes('@') ? raw : `${raw}@pemudamawe.com`;
 const resolved = normalized || ADMIN_EMAIL;
 const {data,error}=await supabase.auth.signInWithPassword({email:resolved,password});
 if(error){document.querySelector('#loginError')?.replaceChildren(document.createTextNode(error.message));toast(error.message,'err');return false;}
 state.session=data.session;toast('Login berhasil');go('admin');return true;
}
init();
