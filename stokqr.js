const DEFAULT_URL="https://script.google.com/macros/s/AKfycbzni2nfsVkpgOOOm2ssWnEDvlvVYq2oE02G1KrSRPkoFTq5AMFkknlFtNhPp1bm1qXW/exec";
let SCRIPT_URL=(function(){try{return localStorage.getItem("stokqr_url")}catch(e){return null}})()||DEFAULT_URL;
const KEY="stokqr_v2",CATS=["Elektronik","ATK","Furnitur","Perkakas"];
const INP="w-full px-3 py-2 rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-900 text-sm";
const BTN="px-4 py-2 rounded-xl font-semibold text-sm bg-slate-900 text-white dark:bg-amber-400 dark:text-slate-900 disabled:opacity-60";
const BTN2="px-4 py-2 rounded-xl font-semibold text-sm bg-slate-100 dark:bg-slate-700";
const CARD="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 p-5";
const TH="p-3 text-left text-xs font-semibold text-slate-500";
const PG_B="px-3 py-2 rounded-xl text-sm font-semibold border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 disabled:opacity-40 disabled:cursor-not-allowed";
const PG_ON="px-3 py-2 rounded-xl text-sm font-semibold bg-amber-400 text-slate-900 border border-amber-400";

let db={items:[],trx:[]},user=null,token=null,pegawai=[],users=[],tab="overview",qr=null,lock=false,fotoB64=null;
try{db=JSON.parse(localStorage.getItem(KEY))||db}catch(e){}
const ss={get:k=>{try{return sessionStorage.getItem(k)}catch(e){return null}},set:(k,v)=>{try{v==null?sessionStorage.removeItem(k):sessionStorage.setItem(k,v)}catch(e){}}};
const save=()=>{try{localStorage.setItem(KEY,JSON.stringify(db))}catch(e){}};
const $=i=>document.getElementById(i);
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const find=k=>db.items.find(i=>i.kode===k);
const admin=()=>user&&user.role==="admin";
const status=i=>i.stok<=0?["Habis","bg-rose-100 text-rose-700"]:i.stok<=i.min?["Menipis","bg-amber-100 text-amber-800"]:["Aman","bg-emerald-100 text-emerald-700"];
const badge=i=>{const s=status(i);return `<span class="px-2 py-0.5 rounded-full text-xs font-semibold ${s[1]}">${s[0]}</span>`};
function toast(m,t="info"){const c={info:"bg-slate-800",success:"bg-emerald-600",error:"bg-rose-600",warning:"bg-amber-600"}[t];const d=document.createElement("div");d.className=`${c} text-white text-sm px-4 py-3 rounded-xl shadow-lg pointer-events-auto`;d.textContent=m;$("toasts").appendChild(d);setTimeout(()=>d.remove(),m.length>60?9000:3200)}

/* ===== PENCARIAN: cari apa saja, banyak kata, urutan bebas ===== */
const norm=s=>String(s??"").toLowerCase();
function cocok(teks,q){
 const w=norm(q).trim().split(/\s+/).filter(Boolean);
 const t=norm(teks);
 return w.every(x=>t.includes(x))}
const teksBrg=i=>Object.entries(i).filter(([k])=>k!=="foto").map(([,v])=>v).concat(status(i)[0]).join(" ");
const teksTrx=r=>{const b=find(r.kode)||{};return Object.values(r).concat([b.nama,b.kat,b.lok]).join(" ")};

/* ===== EXPORT EXCEL (SheetJS) ===== */
function unduhExcel(rows,sheet,nama){
 if(!window.XLSX)return toast("Pustaka Excel belum termuat. Cek koneksi internet.","error");
 if(!rows.length)return toast("Tidak ada data untuk diekspor","warning");
 const ws=XLSX.utils.json_to_sheet(rows);
 ws["!cols"]=Object.keys(rows[0]).map(k=>({wch:Math.min(40,Math.max(k.length,...rows.map(r=>String(r[k]??"").length))+2)}));
 const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,sheet);
 const tgl=new Date().toLocaleDateString("sv-SE",{timeZone:"Asia/Jakarta"});
 XLSX.writeFile(wb,`${nama}-${tgl}.xlsx`);
 toast(`${rows.length} baris diekspor ke Excel`,"success")}

/* ===== SELECT2 ===== */
function s2(el,opt){
 if(!window.jQuery||!jQuery.fn||!jQuery.fn.select2)return;
 const $e=jQuery(el);
 if($e.hasClass("select2-hidden-accessible"))$e.select2("destroy");
 $e.select2(Object.assign({width:"100%",minimumResultsForSearch:6},opt||{}))}
function s2Destroy(root){
 if(!window.jQuery||!jQuery.fn||!jQuery.fn.select2)return;
 jQuery(root).find("select.select2-hidden-accessible").each((_,e)=>{try{jQuery(e).select2("destroy")}catch(x){}})}
function openModal(h){
 s2Destroy("#modal-box");
 $("modal-box").innerHTML=h;$("modal").classList.remove("hidden");$("modal").classList.add("flex");
 if(window.jQuery)jQuery("#modal-box select").each((_,e)=>s2(e,{dropdownParent:jQuery("#modal")}))}
function closeModal(){
 s2Destroy("#modal-box");
 $("modal").classList.add("hidden");$("modal").classList.remove("flex")}

/* Lokasi: diambil dari kolom F sheet Barang (field "lok") */
const lokOpts=cur=>{
 const list=[...new Set([...db.items.map(i=>i.lok),cur].map(v=>String(v||"").trim()).filter(Boolean))]
  .sort((a,b)=>a.localeCompare(b));
 return '<option value="">- Lokasi simpan -</option>'+list.map(n=>`<option value="${esc(n)}" ${n===cur?"selected":""}>${esc(n)}</option>`).join("")};

/* ===== API ===== */
async function api(action,payload={}){
 let d;
 try{
  const r=await fetch(SCRIPT_URL,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({action,payload,token})});
  const t=await r.text();
  try{d=JSON.parse(t)}catch(x){throw new Error("Server tidak mengembalikan JSON. Deploy ulang Web App (New version) dengan akses Anyone, lalu pastikan SCRIPT_URL benar.")}
 }catch(e){if(e instanceof TypeError)throw new Error("Tidak bisa terhubung ke server. Cek SCRIPT_URL dan akses deployment (Who has access = Anyone).");throw e}
 if(d&&d.ok===false){const e=new Error(d.pesan||"Permintaan gagal");e.d=d;if(user&&/Sesi berakhir/.test(e.message))logout(true);throw e}
 return d}
let inflight=null,folderErr="";
function setSync(ok,msg,detail){
 if(ok===false&&!$("sbmsg")&&!$("sync")){toast(detail||msg,"error");return}
 const e=$("sync");if(e){e.textContent=(ok?"● ":"○ ")+msg;e.className="hidden sm:inline text-xs "+(ok?"text-emerald-600":"text-rose-600")}
 const m=$("sbmsg"),bar=$("syncbar");
 if(m){m.textContent=(ok===null?"⟳ ":ok?"● ":"○ ")+(detail||msg)+(ok&&folderErr?" · ⚠ Folder foto: "+folderErr:"");bar.className="px-4 py-2 text-sm text-center "+(ok===null?"bg-slate-200 dark:bg-slate-800":ok?"bg-emerald-100 text-emerald-800":"bg-rose-100 text-rose-800")}}
const busyUI=()=>tab==="scan"||!$("modal").classList.contains("hidden")||/INPUT|SELECT|TEXTAREA/.test((document.activeElement||{}).tagName||"");
function fetchCloudData(manual){
 if(inflight)return inflight;
 inflight=syncRun(manual).then(()=>{inflight=null},()=>{inflight=null});return inflight}
async function getJSON(url){
 let t;
 try{t=await (await fetch(url)).text()}catch(e){throw new Error("Tidak bisa terhubung ke server. Cek SCRIPT_URL, koneksi internet, dan akses deployment (Who has access = Anyone).")}
 try{return JSON.parse(t)}catch(e){throw new Error("Server membalas halaman, bukan data. Biasanya Web App belum di-deploy sebagai Anyone atau belum diberi izin. Klik 'Uji URL' untuk melihat detailnya.")}}
async function syncRun(manual){
 if(manual)setSync(null,"Menyinkronkan…");
 try{
  const d=await getJSON(SCRIPT_URL+"?action=getAll&_="+Date.now());
  if(d.ok===false)throw new Error(d.pesan);
  if(!d.items||!d.trx)throw new Error("Web App masih memakai kode lama. Di Apps Script: Deploy > Manage deployments > Edit > New version > Deploy.");
  const changed=JSON.stringify({items:d.items,trx:d.trx})!==JSON.stringify(db);
  db={items:d.items,trx:d.trx};save();
  const jam=new Date().toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"});
  setSync(true,"Sinkron "+jam,`Tersinkron ${jam} · ${d.items.length} barang · ${d.trx.length} transaksi`);
  if(changed||manual){if(!user)renderLanding();else if(manual||!busyUI())go(tab)}
 }catch(e){setSync(false,"Offline",e.message);if(manual)toast(e.message,"error")}}
async function boot(){
 setSync(null,"Menyinkronkan…");
 if($("sblink"))$("sblink").href=SCRIPT_URL+"?action=ping";
 try{
  const p=await getJSON(SCRIPT_URL+"?action=ping&_="+Date.now());
  if(p.ok===false)throw new Error(p.pesan);
  if(p.versi===undefined)throw new Error("Web App masih memakai kode lama. Di Apps Script: Deploy > Manage deployments > Edit > New version > Deploy.");
  folderErr=p.folderError?(p.folderError+(p.akun?" (akun script: "+p.akun+")":"")):"";
 }catch(e){setSync(false,"Offline",e.message);return}
 await fetchCloudData();
 try{users=(await api("getUsers")).data}catch(e){}}
function ubahUrl(){
 const u=prompt("Tempel URL Web App (berakhiran /exec):",SCRIPT_URL);
 if(u&&/^https:\/\/script\.google\.com\/.+\/exec$/.test(u.trim())){try{localStorage.setItem("stokqr_url",u.trim())}catch(e){}location.reload()}
 else if(u)toast("URL tidak valid","error")}

setInterval(()=>{if(!document.hidden)fetchCloudData()},30000);
document.addEventListener("visibilitychange",()=>{if(!document.hidden)fetchCloudData()});

/* ===== LANDING (katalog 4 per halaman) ===== */
let cPage=1;const C_PER=4;

function renderLanding(){
 const units=db.items.reduce((a,i)=>a+i.stok,0),low=db.items.filter(i=>i.stok<=i.min).length;
 $("lstats").innerHTML=[[db.items.length,"Jenis barang"],[units,"Total unit"],[db.trx.length,"Transaksi tercatat"],[low,"Perlu restok"]].map(s=>`<div><div class="hd text-5xl font-extrabold">${s[0]}</div><div class="text-sm font-medium">${s[1]}</div></div>`).join("");
 const cats=[...new Set([...CATS,...db.items.map(i=>i.kat).filter(Boolean)])],sel=$("c-cat"),cur=sel.value;
 sel.innerHTML='<option value="">Semua kategori</option>'+cats.map(c=>`<option ${c===cur?"selected":""}>${esc(c)}</option>`).join("");
 s2(sel,{width:"12rem"});
 renderCatalog(true)}

/* keep=true: pertahankan halaman (sinkron & tombol halaman). Tanpa argumen: kembali ke halaman 1. */
function renderCatalog(keep){
 if(keep!==true)cPage=1;
 const c=$("c-cat").value;
 const list=db.items.filter(i=>(!c||i.kat===c)&&cocok(teksBrg(i),$("c-q").value));
 const pages=Math.max(1,Math.ceil(list.length/C_PER));
 cPage=Math.min(Math.max(1,cPage),pages);
 const from=(cPage-1)*C_PER,view=list.slice(from,from+C_PER);
 $("catalog").innerHTML=view.length?view.map(i=>`<div class="${CARD}">${i.foto?`<img src="${esc(i.foto)}" alt="" loading="lazy" class="w-full h-32 object-cover rounded-xl mb-3">`:""}<div class="flex justify-between items-start"><span class="text-xs font-mono text-slate-500">${esc(i.kode)}</span>${badge(i)}</div><h3 class="font-semibold mt-2">${esc(i.nama)}</h3><p class="text-sm text-slate-500 dark:text-slate-400">${esc(i.kat)} · ${esc(i.lok)}</p><div class="mt-4 pt-3 border-t border-slate-100 dark:border-slate-800 flex justify-between text-sm"><span>Stok</span><span class="font-bold">${i.stok}</span></div></div>`).join(""):'<p class="col-span-full text-center text-slate-500 py-10">Barang tidak ditemukan.</p>';
 pagerHTML("cpager",list.length,pages,cPage,C_PER,"goCat")}
function goCat(n){cPage=n;renderCatalog(true);$("katalog").scrollIntoView({behavior:"smooth",block:"start"})}

/* Pager bersama: Awal, Sebelumnya, nomor, Berikutnya, Akhir */
function pagerHTML(id,total,pages,cur,per,fn){
 const el=$(id);if(!el)return;
 if(!total||pages<=1){el.innerHTML=total?`<p class="text-sm text-slate-500">Menampilkan ${total} barang</p>`:"";return}
 const start=Math.max(1,Math.min(cur-2,pages-4)),end=Math.min(pages,start+4);
 let nums="";for(let p=start;p<=end;p++)nums+=`<button class="${p===cur?PG_ON:PG_B}" onclick="${fn}(${p})">${p}</button>`;
 const a=(cur-1)*per+1,b=Math.min(cur*per,total);
 el.innerHTML=`<div class="flex flex-wrap justify-center gap-2">
  <button class="${PG_B}" ${cur===1?"disabled":""} onclick="${fn}(1)">&laquo; Awal</button>
  <button class="${PG_B}" ${cur===1?"disabled":""} onclick="${fn}(${cur-1})">&lsaquo; Sebelumnya</button>
  ${nums}
  <button class="${PG_B}" ${cur===pages?"disabled":""} onclick="${fn}(${cur+1})">Berikutnya &rsaquo;</button>
  <button class="${PG_B}" ${cur===pages?"disabled":""} onclick="${fn}(${pages})">Akhir &raquo;</button></div>
  <p class="text-sm text-slate-500">Menampilkan ${a}-${b} dari ${total} barang &middot; Halaman ${cur} / ${pages}</p>`}

function quickCheck(){
 const k=$("quick").value.trim().toUpperCase(),i=find(k);
 $("quick-res").innerHTML=!k?"Hasil pencarian muncul di sini.":i?`<div class="flex justify-between"><b>${esc(i.nama)}</b>${badge(i)}</div><div class="mt-1">Stok <b>${i.stok}</b> unit · ${esc(i.lok)}</div>`:"Kode tidak ditemukan."}

/* ===== AUTH ===== */
async function openLogin(){
 if(!users.length){try{users=(await api("getUsers")).data}catch(e){toast("Gagal memuat daftar pengguna: "+e.message,"error")}}
 openModal(`<h3 class="hd text-3xl font-bold mb-4">Masuk</h3><div class="space-y-3">
 ${users.length?`<select id="lu" class="${INP}">${users.map(u=>`<option value="${esc(u.username)}">${esc(u.nama||u.username)}</option>`).join("")}</select>`:`<input id="lu" placeholder="Username (mis. admin)" class="${INP}"/>`}
 <input id="lp" type="password" placeholder="Password" class="${INP}" onkeydown="if(event.key==='Enter')doLogin()">
 <div class="flex gap-2"><button id="bl" class="${BTN} flex-1" onclick="doLogin()">Masuk ke dashboard</button><button class="${BTN2}" onclick="closeModal()">Batal</button></div></div>`)}
async function doLogin(){
 const u=$("lu").value,b=$("bl");b.disabled=true;
 try{const r=await api("login",{username:u,password:$("lp").value});token=r.token;user=r.user;ss.set("tk",token);closeModal();toast("Selamat datang, "+user.nama,"success");enter()}
 catch(e){b.disabled=false;if(e.d&&e.d.perluBuatPassword)return openBuat(u);toast(e.message,"error")}}
function openBuat(u){openModal(`<h3 class="hd text-3xl font-bold mb-1">Buat password</h3><p class="text-sm text-slate-500 mb-4">Akun <b>${esc(u)}</b> belum punya password.</p><div class="space-y-3">
 <input id="n1" type="password" placeholder="Password baru (min. 6 karakter)" class="${INP}"><input id="n2" type="password" placeholder="Ulangi password" class="${INP}">
 <div class="flex gap-2"><button class="${BTN} flex-1" onclick="doBuat('${esc(u)}')">Simpan dan masuk</button><button class="${BTN2}" onclick="closeModal()">Batal</button></div></div>`)}
async function doBuat(u){
 const p=$("n1").value;if(p!==$("n2").value)return toast("Password tidak sama","error");
 try{await api("buatPassword",{username:u,password:p});const r=await api("login",{username:u,password:p});token=r.token;user=r.user;ss.set("tk",token);closeModal();enter()}catch(e){toast(e.message,"error")}}
function openGanti(){openModal(`<h3 class="hd text-3xl font-bold mb-4">Ganti password</h3><div class="space-y-3">
 <input id="g1" type="password" placeholder="Password lama" class="${INP}"><input id="g2" type="password" placeholder="Password baru (min. 6)" class="${INP}">
 <div class="flex gap-2"><button class="${BTN} flex-1" onclick="doGanti()">Simpan</button><button class="${BTN2}" onclick="closeModal()">Batal</button></div></div>`)}
async function doGanti(){try{await api("gantiPassword",{lama:$("g1").value,baru:$("g2").value});closeModal();toast("Password diganti","success")}catch(e){toast(e.message,"error")}}
function logout(silent){
 if(token&&!silent)api("logout").catch(()=>{});
 user=null;token=null;ss.set("tk",null);stopScan();$("app").classList.add("hidden");$("landing").classList.remove("hidden");renderLanding();if(silent)toast("Sesi berakhir, silakan masuk lagi","warning")}
async function enter(){
 $("landing").classList.add("hidden");$("app").classList.remove("hidden");
 $("uname").textContent=user.nama;$("role").textContent=admin()?"Admin":"Staf";go("overview");
 try{pegawai=(await api("getPegawai")).data}catch(e){}}
const pegOpts=()=>'<option value="">- Pegawai -</option>'+pegawai.map(p=>`<option>${esc(p.nama)}</option>`).join("");

/* ===== TABS ===== */
function go(t){
 tab=t;if(t!=="scan")stopScan();
 const T=[["overview","Ringkasan"],["barang","Barang"],["trx","Transaksi"],["scan","Scan QR"]];
 $("tabs").innerHTML=T.map(x=>`<button onclick="go('${x[0]}')" class="px-4 py-2.5 text-sm font-semibold whitespace-nowrap border-b-2 ${x[0]===t?"border-amber-400":"border-transparent text-slate-500"}">${x[1]}</button>`).join("");
 ({overview:vOverview,barang:vBarang,trx:vTrx,scan:vScan})[t]()}

function vOverview(){
 const units=db.items.reduce((a,i)=>a+i.stok,0),low=db.items.filter(i=>i.stok<=i.min),td=db.trx.filter(t=>t.tgl===new Date().toLocaleDateString("sv-SE",{timeZone:"Asia/Jakarta"})).length;
 const st=[["Jenis barang",db.items.length],["Total unit",units],["Stok menipis",low.length],["Transaksi hari ini",td]];
 $("view").innerHTML=`<div class="grid grid-cols-2 lg:grid-cols-4 gap-4">${st.map(s=>`<div class="${CARD}"><div class="text-xs text-slate-500">${s[0]}</div><div class="hd text-4xl font-extrabold mt-1">${s[1]}</div></div>`).join("")}</div>
 <div class="grid lg:grid-cols-2 gap-4 mt-4">
 <div class="${CARD}"><h3 class="hd text-2xl font-bold mb-3">Perlu restok</h3>${low.length?low.map(i=>`<div class="flex justify-between py-2 border-b border-slate-100 dark:border-slate-800 text-sm"><span>${esc(i.nama)}</span><span>${i.stok} / min ${i.min} ${badge(i)}</span></div>`).join(""):'<p class="text-sm text-slate-500">Semua stok aman.</p>'}</div>
 <div class="${CARD}"><h3 class="hd text-2xl font-bold mb-3">Transaksi terbaru</h3>${db.trx.slice(0,5).map(t=>`<div class="flex justify-between py-2 border-b border-slate-100 dark:border-slate-800 text-sm"><span>${esc(find(t.kode)?.nama||t.kode)}</span><span class="${t.jenis==="Masuk"?"text-emerald-600":"text-rose-600"} font-semibold">${t.jenis==="Masuk"?"+":"-"}${t.jml}</span></div>`).join("")||'<p class="text-sm text-slate-500">Belum ada transaksi.</p>'}</div></div>
 <div class="mt-4 flex gap-2"><button class="${BTN}" onclick="go('scan')">Buka scanner</button><button class="${BTN2}" onclick="openTrx()">Catat transaksi manual</button></div>`}

/* ----- Tab Barang: cari apa saja, 5 per halaman, export Excel ----- */
let bPage=1,bQ="";const B_PER=5;

function getBrg(){return db.items.filter(i=>cocok(teksBrg(i),bQ))}

function exportBrg(){
 unduhExcel(getBrg().map(i=>({"Kode":i.kode,"Nama":i.nama,"Kategori":i.kat,"Lokasi":i.lok,"Stok":Number(i.stok)||0,"Stok Minimum":Number(i.min)||0,"Status":status(i)[0]})),"Barang","barang-stokqr")}

function vBarang(){
 $("view").innerHTML=`<div class="flex flex-col sm:flex-row justify-between sm:items-end gap-3 mb-4"><div class="w-full sm:max-w-xs"><label for="b-q" class="block text-xs text-slate-500 mb-1">Cari barang</label><input id="b-q" value="${esc(bQ)}" oninput="bQ=this.value;rowsBarang()" placeholder="Cari apa saja: nama, kode, kategori, lokasi, status" class="${INP}"></div><div class="flex flex-wrap gap-2"><button class="${BTN2}" onclick="exportBrg()">📥 Export Excel</button><button class="${BTN2}" onclick="printLabels(db.items)">Cetak semua label</button><button class="${BTN}" onclick="openItem()">Data barang baru</button></div></div>
 <div class="${CARD} overflow-x-auto p-2"><table class="w-full text-sm"><thead><tr><th class="${TH}">Foto</th><th class="${TH}">Kode</th><th class="${TH}">Nama</th><th class="${TH}">Kategori</th><th class="${TH}">Lokasi</th><th class="${TH}">Stok</th><th class="${TH}">Status</th><th class="${TH} text-right">Aksi</th></tr></thead><tbody id="b-body" class="divide-y divide-slate-100 dark:divide-slate-800"></tbody></table></div>
 <div class="mt-5 flex flex-col items-center gap-3" id="b-pager"></div>`;
 rowsBarang(true)}

function rowsBarang(keep){
 if(keep!==true)bPage=1;
 const list=getBrg();
 const pages=Math.max(1,Math.ceil(list.length/B_PER));
 bPage=Math.min(Math.max(1,bPage),pages);
 const from=(bPage-1)*B_PER,view=list.slice(from,from+B_PER);
 $("b-body").innerHTML=view.length?view.map(i=>`<tr><td class="p-3">${i.foto?`<img src="${esc(i.foto)}" alt="" class="w-10 h-10 rounded-lg object-cover">`:"–"}</td><td class="p-3 font-mono text-xs">${esc(i.kode)}</td><td class="p-3 font-semibold">${esc(i.nama)}</td><td class="p-3">${esc(i.kat)}</td><td class="p-3">${esc(i.lok)}</td><td class="p-3 font-bold">${i.stok}<span class="text-xs text-slate-400 font-normal"> / min ${i.min}</span></td><td class="p-3">${badge(i)}</td><td class="p-3 text-right whitespace-nowrap"><button title="Label QR" class="px-2 py-1" onclick="openQR('${esc(i.kode)}')">🏷️</button>${admin()?`<button title="Ubah" class="px-2 py-1" onclick="openItem('${esc(i.kode)}')">✏️</button><button title="Hapus" class="px-2 py-1" onclick="delItem('${esc(i.kode)}')">🗑️</button>`:""}</td></tr>`).join(""):`<tr><td colspan="8" class="p-8 text-center text-slate-500">Barang tidak ditemukan.</td></tr>`;
 pagerHTML("b-pager",list.length,pages,bPage,B_PER,"goBrg")}
function goBrg(n){bPage=n;rowsBarang(true)}

/* Foto: perkecil di browser sebelum diunggah */
function kecilkan(file,max=1200){return new Promise((ok,err)=>{const img=new Image(),u=URL.createObjectURL(file);
 img.onload=()=>{const s=Math.min(1,max/Math.max(img.width,img.height)),c=document.createElement("canvas");c.width=img.width*s;c.height=img.height*s;c.getContext("2d").drawImage(img,0,0,c.width,c.height);URL.revokeObjectURL(u);const d=c.toDataURL("image/jpeg",.8);ok({b64:d.split(",")[1],url:d})};
 img.onerror=()=>err(new Error("Gambar tidak dapat dibaca"));img.src=u})}
async function onFoto(e){const f=e.target.files[0],p=$("f-prev");fotoB64=null;if(!f){p.classList.add("hidden");return}
 try{const r=await kecilkan(f);fotoB64={b64:r.b64,nama:f.name.replace(/\.\w+$/,"")+".jpg"};p.src=r.url;p.classList.remove("hidden")}catch(x){toast(x.message,"error")}}

async function tesFolder(){
 try{const r=await api("tesFolder");alert("Folder foto OK: "+r.folder+"\nAkun script: "+r.akun+"\nBisa menulis file.")}
 catch(e){alert("Tes folder GAGAL:\n\n"+e.message)}}
function openItem(k){
 fotoB64=null;const i=k?find(k):{kode:"",nama:"",kat:CATS[0],stok:0,min:5,lok:"",foto:""};
 openModal(`<h3 class="hd text-3xl font-bold mb-4">${k?"Ubah barang":"Data barang baru"}</h3><div class="space-y-3">
 <input id="f-kode" value="${esc(i.kode)}" ${k?"disabled":""} placeholder="Kode barang (kosong = otomatis)" class="${INP}"><input id="f-nama" value="${esc(i.nama)}" placeholder="Nama barang" class="${INP}">
 <div class="grid grid-cols-2 gap-3">
  <select id="f-kat" class="${INP}">
   ${CATS.map(c=>`<option value="${esc(c)}" ${c===i.kat?"selected":""}>${esc(c)}</option>`).join("")}
  </select>
  <select id="f-lok" class="${INP}">${lokOpts(i.lok)}</select>
 </div>
 <div class="grid grid-cols-2 gap-3"><label class="text-xs text-slate-500">Stok<input id="f-stok" type="number" min="0" value="${i.stok}" class="${INP}"></label><label class="text-xs text-slate-500">Stok minimum<input id="f-min" type="number" min="0" value="${i.min}" class="${INP}"></label></div>
 <select id="f-peg" class="${INP}">${pegOpts()}</select>
 <label class="text-xs text-slate-500">Foto barang<input id="f-foto" type="file" accept="image/*" capture="environment" onchange="onFoto(event)" class="${INP}"></label>
 <img id="f-prev" alt="" class="hidden w-32 h-32 object-cover rounded-xl"/><button type="button" class="text-xs underline text-slate-500" onclick="tesFolder()">Tes folder foto</button>
 <div class="flex gap-2"><button id="f-ok" class="${BTN} flex-1" onclick="saveItem('${k||""}')">Simpan ke Sheets</button><button class="${BTN2}" onclick="closeModal()">Batal</button></div></div>`)}

async function saveItem(k){
 const d={kode:k||$("f-kode").value.trim().toUpperCase(),nama:$("f-nama").value.trim(),kat:$("f-kat").value.trim(),lok:$("f-lok").value.trim(),stok:Math.max(0,+$("f-stok").value||0),min:Math.max(0,+$("f-min").value||0),pegawai:$("f-peg").value,foto:k?find(k).foto:""};
 if(!d.nama)return toast("Nama barang wajib diisi","warning");
 const b=$("f-ok");b.disabled=true;b.textContent="Menyimpan…";
 try{
  if(fotoB64){try{const up=await api("uploadGambar",{base64:fotoB64.b64,mimeType:"image/jpeg",namaFile:fotoB64.nama});d.foto=up.thumb}catch(e){if(/Sesi berakhir/.test(e.message))throw e;alert("Foto GAGAL diunggah (data barang tetap tersimpan):\n\n"+e.message)}}
  const r=await api(k?"saveItem":"simpanData",d);
  closeModal();toast(`Barang ${r.kode} tersimpan`,"success");await fetchCloudData();go("barang")
 }catch(e){b.disabled=false;b.textContent="Simpan ke Sheets";toast(e.message,"error")}}

async function delItem(k){
 if(!confirm("Hapus barang "+k+"?"))return;
 try{await api("delItem",{kode:k});toast("Barang dihapus");await fetchCloudData();go("barang")}catch(e){toast(e.message,"error")}}

/* ----- Tab Transaksi: cari apa saja, filter tanggal, export Excel ----- */
let tQ="",tFrom="",tTo="";

function getTrx(){
 let f=tFrom,t=tTo;
 if(f&&t&&f>t){const x=f;f=t;t=x}
 return db.trx.filter(r=>{
  const d=String(r.tgl||"").slice(0,10);
  if(f&&d<f)return false;
  if(t&&d>t)return false;
  return cocok(teksTrx(r),tQ)})}

function exportTrx(){
 unduhExcel(getTrx().map(r=>({"ID":r.id,"Tanggal":r.tgl,"Kode":r.kode,"Nama Barang":find(r.kode)?.nama||"","Jenis":r.jenis,"Jumlah":Number(r.jml)||0,"Pegawai":r.pegawai||"","Keterangan":r.ket||""})),"Transaksi","transaksi-stokqr")}

function vTrx(){
 $("view").innerHTML=`<div class="flex flex-col sm:flex-row justify-between sm:items-center gap-3 mb-4"><h2 class="hd text-3xl font-bold">Riwayat transaksi</h2><div class="flex gap-2"><button class="${BTN2}" onclick="exportTrx()">📥 Export Excel</button><button class="${BTN}" onclick="openTrx()">Catat transaksi</button></div></div>
 <div class="flex flex-wrap items-end gap-3 mb-4">
  <div class="w-full sm:w-72"><label for="t-q" class="block text-xs text-slate-500 mb-1">Cari transaksi</label><input id="t-q" value="${esc(tQ)}" oninput="tQ=this.value;rowsTrx()" placeholder="Cari apa saja: barang, kode, pegawai, jenis, tanggal" class="${INP}"></div>
  <div class="w-[48%] sm:w-44"><label for="t-from" class="block text-xs text-slate-500 mb-1">Dari tanggal</label><input id="t-from" type="date" value="${esc(tFrom)}" oninput="tFrom=this.value;rowsTrx()" class="${INP}"></div>
  <div class="w-[48%] sm:w-44"><label for="t-to" class="block text-xs text-slate-500 mb-1">Sampai tanggal</label><input id="t-to" type="date" value="${esc(tTo)}" oninput="tTo=this.value;rowsTrx()" class="${INP}"></div>
  <button class="${BTN2}" onclick="resetTrx()">Reset</button>
 </div>
 <div class="${CARD} overflow-x-auto p-2"><table class="w-full text-sm"><thead><tr><th class="${TH}">ID</th><th class="${TH}">Tanggal</th><th class="${TH}">Barang</th><th class="${TH}">Jenis</th><th class="${TH}">Jumlah</th><th class="${TH}">Pegawai</th><th class="${TH}">Keterangan</th></tr></thead><tbody id="t-body" class="divide-y divide-slate-100 dark:divide-slate-800"></tbody></table></div>
 <p class="mt-3 text-sm text-slate-500" id="t-info"></p>`;
 rowsTrx()}

function resetTrx(){tQ="";tFrom="";tTo="";vTrx()}

function rowsTrx(){
 const list=getTrx();
 $("t-body").innerHTML=list.length?list.map(r=>`<tr><td class="p-3 font-mono text-xs">${esc(r.id)}</td><td class="p-3">${esc(r.tgl)}</td><td class="p-3 font-semibold">${esc(find(r.kode)?.nama||r.kode)}</td><td class="p-3"><span class="px-2 py-0.5 rounded-full text-xs font-semibold ${r.jenis==="Masuk"?"bg-emerald-100 text-emerald-700":"bg-rose-100 text-rose-700"}">${esc(r.jenis)}</span></td><td class="p-3 font-bold">${r.jml}</td><td class="p-3">${esc(r.pegawai||"-")}</td><td class="p-3 text-slate-500">${esc(r.ket||"-")}</td></tr>`).join(""):`<tr><td colspan="7" class="p-8 text-center text-slate-500">Transaksi tidak ditemukan.</td></tr>`;
 const info=$("t-info");if(info)info.textContent=`Menampilkan ${list.length} dari ${db.trx.length} transaksi`}

function openTrx(k){
 openModal(`<h3 class="hd text-3xl font-bold mb-4">Catat transaksi</h3><div class="space-y-3">
 <select id="t-kode" class="${INP}">${db.items.map(i=>`<option value="${esc(i.kode)}" ${i.kode===k?"selected":""}>${esc(i.kode)} - ${esc(i.nama)} (stok ${i.stok})</option>`).join("")}</select>
 <div class="grid grid-cols-2 gap-3"><select id="t-jenis" class="${INP}"><option>Masuk</option><option>Keluar</option></select><input id="t-jml" type="number" min="1" value="1" class="${INP}"></div>
 <select id="t-peg" class="${INP}">${pegOpts()}</select><input id="t-ket" placeholder="Keterangan (opsional)" class="${INP}">
 <div class="flex gap-2"><button id="t-ok" class="${BTN} flex-1" onclick="saveTrx()">Simpan transaksi</button><button class="${BTN2}" onclick="cancelTrx()">Batal</button></div></div>`)}
function cancelTrx(){closeModal();lock=false;if(tab==="scan")startScan()}
async function saveTrx(){
 const i=find($("t-kode").value),jenis=$("t-jenis").value,jml=Math.floor(+$("t-jml").value);
 if(!jml||jml<1)return toast("Jumlah minimal 1","error");
 if(jenis==="Keluar"&&jml>i.stok)return toast(`Stok ${i.nama} hanya ${i.stok}`,"error");
 const b=$("t-ok");b.disabled=true;
 try{
  const r=await api("saveTrx",{trx:{kode:i.kode,jenis,jml,ket:$("t-ket").value.trim(),pegawai:$("t-peg").value}});
  i.stok=r.stok;db.trx.unshift(r.trx);save();
  closeModal();lock=false;toast(`${jenis} ${jml} unit ${i.nama} tercatat`,"success");
  if(tab==="scan")startScan();else go(tab);
 }catch(e){b.disabled=false;toast(e.message,"error")}}

/* ===== SCANNER ===== */
function vScan(){
 $("view").innerHTML=`<div class="${CARD} max-w-xl mx-auto"><h2 class="hd text-3xl font-bold">Scan QR barang</h2><p class="text-sm text-slate-500 mb-4">Arahkan kamera ke label QR. Form transaksi terbuka otomatis.</p>
 <div id="reader" class="rounded-xl bg-slate-900 min-h-[240px] overflow-hidden"></div>
 <div class="flex gap-2 mt-3"><button class="${BTN}" onclick="startScan()">Aktifkan kamera</button><button class="${BTN2}" onclick="stopScan()">Matikan</button></div>
 <div class="mt-6 pt-4 border-t border-slate-100 dark:border-slate-800"><label class="text-xs text-slate-500">Tanpa kamera? Masukkan kode barang</label><div class="flex gap-2 mt-1"><input id="sim" placeholder="BRG-001" class="${INP}"><button class="${BTN2}" onclick="onScan($('sim').value)">Proses</button></div></div></div>`}
async function startScan(){
 if(!window.Html5Qrcode)return toast("Pustaka kamera belum termuat. Gunakan input kode.","warning");
 await stopScan();if(!$("reader"))return;
 qr=new Html5Qrcode("reader");
 try{await qr.start({facingMode:"environment"},{fps:10,qrbox:220},onScan,()=>{})}catch(e){qr=null;toast("Kamera tidak dapat diakses. Cek izin browser.","error")}}
async function stopScan(){if(qr){try{await qr.stop();qr.clear()}catch(e){}qr=null}}
async function onScan(code){
 if(lock)return;code=String(code).trim().toUpperCase();
 if(!find(code)){lock=true;toast(`Kode ${code} tidak terdaftar`,"error");setTimeout(()=>lock=false,1500);return}
 lock=true;await stopScan();openTrx(code)}

/* ===== LABEL QR ===== */
function openQR(k){const i=find(k);
 openModal(`<div class="text-center"><p class="text-xs text-slate-500">Label QR</p><h3 class="font-bold text-lg">${esc(i.nama)}</h3><div id="qr1" class="my-4 flex justify-center bg-white p-3 rounded-xl w-fit mx-auto"></div><p class="font-mono text-sm font-bold">${esc(i.kode)}</p>
 <div class="flex gap-2 mt-5"><button class="${BTN} flex-1" onclick="printLabels([find('${esc(i.kode)}')])">Cetak label</button><button class="${BTN2}" onclick="closeModal()">Tutup</button></div></div>`);
 setTimeout(()=>new QRCode($("qr1"),{text:i.kode,width:160,height:160}),30)}
function printLabels(list){
 if(!list.length)return toast("Belum ada barang","warning");
 $("print-area").innerHTML=list.map((i,n)=>`<div style="border:1px solid #999;border-radius:8px;padding:8px;text-align:center;break-inside:avoid"><div style="font-size:11px;font-weight:700">${esc(i.nama)}</div><div id="pq${n}" style="display:flex;justify-content:center;margin:6px 0"></div><div style="font:12px monospace">${esc(i.kode)}</div></div>`).join("");
 list.forEach((i,n)=>new QRCode($("pq"+n),{text:i.kode,width:90,height:90}));
 setTimeout(()=>window.print(),250)}

/* ===== INIT ===== */
renderLanding();
boot();
(async()=>{const t=ss.get("tk");if(!t)return;token=t;
 try{const r=await api("cekSesi");if(r.ok){user=r.user;enter()}}catch(e){token=null;ss.set("tk",null)}})();

//]]>