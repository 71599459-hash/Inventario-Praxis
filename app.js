
(() => {
  const STATE_KEY = "praxis-state-v2";
  const PAGE_SIZE = 30;
  let state = null;
  let currentView = "dashboard";
  let page = 1;
  let selectedId = null;

  const $ = s => document.querySelector(s);
  const $$ = s => [...document.querySelectorAll(s)];
  const esc = value => String(value ?? "").replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
  const norm = value => PraxisExcel.norm(value);
  const today = () => new Date().toISOString().slice(0,10);

  const statusClass = s => {
    const x = norm(s);
    if (x.includes("OPERAT")) return "ok";
    if (x.includes("MANT") || x.includes("REGULAR")) return "warn";
    if (x.includes("AVERI") || x.includes("BAJA") || x.includes("MALO")) return "bad";
    return "info";
  };

  const locationLabel = r => r.locationType === "ALMACEN" ? "Almacén" : r.locationType === "ASIGNADO" ? "Asignado" : "Sede / área";

  function toast(message, type="") {
    const el = document.createElement("div");
    el.className = `toast ${type}`;
    el.textContent = message;
    $("#toastHost").appendChild(el);
    setTimeout(() => el.remove(), 3500);
  }

  async function persist() {
    await PraxisDB.set(STATE_KEY, state);
  }

  function groupCount(arr, getter) {
    const out = {};
    arr.forEach(item => {
      const key = typeof getter === "function" ? getter(item) : item[getter];
      if (!key) return;
      out[key] = (out[key] || 0) + 1;
    });
    return out;
  }

  const topEntries = (obj, limit=10) => Object.entries(obj).sort((a,b)=>b[1]-a[1]).slice(0,limit);

  function activeInventory() {
    return state?.inventory || [];
  }

  function collaborators() {
    const map = new Map();
    activeInventory().filter(r=>r.responsable).forEach(r => {
      const key = `${norm(r.responsable)}|${String(r.dni||"").trim()}`;
      if (!map.has(key)) map.set(key,{ key, name:r.responsable, dni:r.dni, items:[], sites:new Set(), areas:new Set() });
      const p = map.get(key); p.items.push(r);
      if (r.sede) p.sites.add(r.sede); if (r.area) p.areas.add(r.area);
    });
    return [...map.values()].sort((a,b)=>a.name.localeCompare(b.name,"es"));
  }

  function setView(view) {
    currentView = view;
    $$(".view").forEach(v => v.classList.toggle("active", v.dataset.viewPanel === view));
    $$(".nav-item").forEach(b => b.classList.toggle("active", b.dataset.view === view));
    const titles = {
      dashboard:["Inventario TIC","Gestión de equipos tecnológicos"],
      inventory:["Inventario maestro","Búsqueda, edición y control del Código TIC"],
      general:["General / Administrativos","Equipos a cargo del personal administrativo"],
      sites:["Sedes","Inventario separado por cada sede"],
      people:["Colaboradores","Equipos y componentes a cargo"],
      warehouse:["Almacén TIC","Control de equipos almacenados"],
      networks:["Cámaras y redes","Infraestructura tecnológica y seguridad"],
      materials:["Materiales","Insumos, materiales y herramientas"],
      movements:["Movimientos","Transferencias e historial"],
      reports:["Reportes","Resumen del inventario"]
    };
    $("#pageTitle").textContent = titles[view][0];
    $("#pageSubtitle").textContent = titles[view][1];
    if (window.innerWidth < 850) $("#sidebar").classList.remove("open");
  }

  function kpiCard(kind, number, label) {
    return `<div class="kpi ${kind}"><div class="number">${number}</div><div class="label">${esc(label)}</div></div>`;
  }

  function renderBars(target, entries) {
    const max = Math.max(1,...entries.map(x=>x[1]));
    $(target).innerHTML = entries.map(([k,v]) => `
      <div class="bar-row">
        <span title="${esc(k)}">${esc(k.length>18?k.slice(0,18)+"…":k)}</span>
        <div class="bar-track"><div class="bar-fill" style="width:${Math.max(3,v/max*100)}%"></div></div>
        <b>${v}</b>
      </div>`).join("") || `<div class="muted">Sin datos.</div>`;
  }

  function renderDashboard() {
    const inv = activeInventory();
    const coded = inv.filter(r=>r.codigo).length;
    const assigned = inv.filter(r=>r.locationType==="ASIGNADO").length;
    const warehouse = inv.filter(r=>r.locationType==="ALMACEN").length;
    const maintenance = inv.filter(r=>norm(r.estado).includes("MANT")).length;
    const noCode = inv.filter(r=>!r.codigo).length;
    const people = collaborators().length;
    $("#kpiGrid").innerHTML = [
      kpiCard("blue", coded, "Con Código TIC"),
      kpiCard("green", assigned, "Asignados a personal"),
      kpiCard("purple", warehouse, "En almacén TIC"),
      kpiCard("orange", people, "Colaboradores"),
      kpiCard("red", maintenance, "En mantenimiento"),
      kpiCard("cyan", noCode, "Sin Código TIC")
    ].join("");

    renderBars("#categoryBars", topEntries(groupCount(inv.filter(r=>r.equipo),"equipo"),10));
    renderBars("#siteBars", topEntries(groupCount(inv.filter(r=>r.sede),"sede"),9));

    const q = [
      [inv.length,"Registros maestros"],
      [state.stats?.duplicateCodes||0,"Códigos consolidados"],
      [inv.filter(r=>r.serie).length,"Con serie/código"],
      [state.transactions?.length||0,"Registros históricos"]
    ];
    $("#qualityGrid").innerHTML = q.map(([n,l])=>`<div class="quality-box"><b>${n}</b><span>${esc(l)}</span></div>`).join("");
    $("#dashboardTable").innerHTML = tableHtml(inv.slice(0,10), false);
  }

  function fillFilters() {
    const setOptions = (selector, values, firstText) => {
      const el = $(selector); const current = el.value;
      const opts = [...new Set(values.filter(Boolean))].sort((a,b)=>a.localeCompare(b,"es"));
      el.innerHTML = `<option value="">${esc(firstText)}</option>` + opts.map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join("");
      if (opts.includes(current)) el.value = current;
    };
    setOptions("#siteFilter", activeInventory().map(r=>r.sede), "Todas las sedes");
    setOptions("#categoryFilter", activeInventory().map(r=>r.equipo), "Todas las categorías");
    setOptions("#statusFilter", activeInventory().map(r=>r.estado), "Todos los estados");
  }

  function inventoryFiltered() {
    const q = norm($("#inventorySearch").value);
    const site = $("#siteFilter").value, category=$("#categoryFilter").value, status=$("#statusFilter").value, location=$("#locationFilter").value;
    return activeInventory().filter(r => {
      const hay = norm([r.codigo,r.codigoPadre,r.equipo,r.descripcion,r.marca,r.modelo,r.serie,r.responsable,r.dni,r.sede,r.area,r.estado,r.condicion,r.situacion,r.observaciones].join(" "));
      return (!q || hay.includes(q)) && (!site||r.sede===site) && (!category||r.equipo===category) && (!status||r.estado===status) && (!location||r.locationType===location);
    });
  }

  function tableHtml(records, paginate=false) {
    let items=records,total=records.length,start=0,pages=1;
    if (paginate) {
      pages=Math.max(1,Math.ceil(total/PAGE_SIZE)); page=Math.min(page,pages); start=(page-1)*PAGE_SIZE;
      items=records.slice(start,start+PAGE_SIZE);
    }
    const rows = items.map(r=>`
      <tr>
        <td><span class="code">${esc(r.codigo||r.id)}</span>${r.codigoPadre?`<div class="muted">Padre: ${esc(r.codigoPadre)}</div>`:""}</td>
        <td><b>${esc(r.equipo||"Sin tipo")}</b>${r.detalle?`<div class="muted">${esc(r.detalle)}</div>`:""}</td>
        <td>${esc(r.marca)}${r.modelo?`<div class="muted">${esc(r.modelo)}</div>`:""}</td>
        <td>${esc(r.serie)}</td>
        <td>${esc(r.sede)}</td>
        <td>${esc(r.area)}</td>
        <td>${esc(r.responsable)}${r.dni?`<div class="muted">DNI: ${esc(r.dni)}</div>`:""}</td>
        <td><span class="badge ${statusClass(r.estado)}">${esc(r.estado||"SIN ESTADO")}</span><div class="muted">${locationLabel(r)}</div></td>
        <td><div class="table-actions"><button data-open="${esc(r.id)}">Ver</button></div></td>
      </tr>`).join("");

    let html=`<div class="table-wrap"><table><thead><tr><th>Código TIC</th><th>Equipo / material</th><th>Marca / modelo</th><th>Serie</th><th>Sede</th><th>Área</th><th>Responsable</th><th>Estado</th><th></th></tr></thead><tbody>${rows||`<tr><td colspan="9">No hay registros.</td></tr>`}</tbody></table></div>`;
    if (paginate) {
      html += `<div class="pagination"><span>Mostrando ${total?start+1:0}-${Math.min(start+PAGE_SIZE,total)} de ${total}</span><div class="page-buttons">
        <button data-page="${Math.max(1,page-1)}">‹</button><button class="active">${page}</button><button data-page="${Math.min(pages,page+1)}">›</button>
      </div></div>`;
    }
    return html;
  }

  function wireTables() {
    $$("[data-open]").forEach(b=>b.onclick=()=>openAsset(b.dataset.open));
    $$("[data-page]").forEach(b=>b.onclick=()=>{page=Number(b.dataset.page);renderInventory()});
  }

  function renderInventory() {
    $("#inventoryTable").innerHTML = tableHtml(inventoryFiltered(), true); wireTables();
  }

  function renderGeneral() {
    const rows=activeInventory().filter(r=>r.source==="General" || (r.responsable && !["TIC - Almacén","Cámaras y Redes","Insum, mat y herra"].includes(r.source)));
    $("#generalTable").innerHTML=tableHtml(rows,false); wireTables();
  }

  function renderWarehouse() {
    const rows=activeInventory().filter(r=>r.locationType==="ALMACEN");
    $("#warehouseTable").innerHTML=tableHtml(rows,false); wireTables();
  }

  function renderNetworks() {
    const re=/CAMARA|NVR|DVR|SWITCH|ROUTER|ACCESS POINT|AP\b|RED|CCTV|BALUN|RACK/;
    const rows=activeInventory().filter(r=>r.source==="Cámaras y Redes" || re.test(norm(`${r.equipo} ${r.descripcion} ${r.detalle}`)));
    $("#networkTable").innerHTML=tableHtml(rows,false); wireTables();
  }

  function renderMaterials() {
    const rows=activeInventory().filter(r=>r.source==="Insum, mat y herra");
    $("#materialTable").innerHTML=tableHtml(rows,false); wireTables();
  }

  function renderSites() {
    const inv=activeInventory(), groups=groupCount(inv.filter(r=>r.sede),"sede");
    $("#siteCards").innerHTML=topEntries(groups,50).map(([site,n])=>{
      const coded=inv.filter(r=>r.sede===site&&r.codigo).length;
      const wh=inv.filter(r=>r.sede===site&&r.locationType==="ALMACEN").length;
      const assigned=inv.filter(r=>r.sede===site&&r.locationType==="ASIGNADO").length;
      return `<article class="site-card" data-site="${esc(site)}"><h4>${esc(site)}</h4><div class="card-number">${n}</div><div class="card-meta">${coded} con Código TIC · ${assigned} asignados · ${wh} almacén</div></article>`;
    }).join("");
    $$("[data-site]").forEach(c=>c.onclick=()=>{
      setView("inventory");$("#siteFilter").value=c.dataset.site;$("#inventorySearch").value="";page=1;renderInventory();
    });
  }

  function renderPeople() {
    const q=norm($("#peopleSearch").value);
    const ps=collaborators().filter(p=>!q||norm([p.name,p.dni,...p.sites,...p.areas].join(" ")).includes(q));
    $("#peopleCards").innerHTML=ps.map(p=>`
      <article class="person-card" data-person="${esc(p.key)}">
        <h4>${esc(p.name)}</h4><div class="card-number">${p.items.length}</div>
        <div class="card-meta">equipos / registros a cargo<br>DNI: ${esc(p.dni||"—")}<br>${esc([...p.sites].join(", "))}</div>
      </article>`).join("") || `<div class="muted">No se encontraron colaboradores.</div>`;
    $$("[data-person]").forEach(c=>c.onclick=()=>openPerson(c.dataset.person));
  }

  function renderMovements() {
    const web = (state.webMovements||[]).map(m=>({...m,kind:"WEB"}));
    const tx = (state.transactions||[]).map(t=>({...t,kind:"EXCEL"}));
    const items=[...web,...tx].sort((a,b)=>String(b.fecha||"").localeCompare(String(a.fecha||"")));
    const rows=items.slice(0,500).map(m=>`<tr><td>${esc(m.fecha)}</td><td>${esc(m.kind)}</td><td>${esc(m.source||m.type||"")}</td><td><span class="code">${esc(m.codigo||"")}</span></td><td>${esc(m.equipo||"")}</td><td>${esc(m.from||"")}</td><td>${esc(m.to||[m.sede,m.area].filter(Boolean).join(" / "))}</td><td>${esc(m.responsable||"")}</td><td>${esc(m.observaciones||m.motivo||"")}</td></tr>`).join("");
    $("#movementTable").innerHTML=`<div class="table-wrap"><table><thead><tr><th>Fecha</th><th>Tipo</th><th>Origen</th><th>Código TIC</th><th>Equipo</th><th>Desde</th><th>Hacia</th><th>Responsable</th><th>Observación</th></tr></thead><tbody>${rows||`<tr><td colspan="9">Sin movimientos.</td></tr>`}</tbody></table></div>`;
  }

  function reportList(target,obj) {
    $(target).innerHTML=`<div class="report-list">${topEntries(obj,30).map(([k,v])=>`<div class="report-row"><span>${esc(k)}</span><b>${v}</b></div>`).join("")}</div>`;
  }

  function renderReports() {
    const inv=activeInventory();
    reportList("#reportSites",groupCount(inv.filter(r=>r.sede),"sede"));
    reportList("#reportCategories",groupCount(inv.filter(r=>r.equipo),"equipo"));
    reportList("#reportStatuses",groupCount(inv.filter(r=>r.estado),"estado"));
  }

  function renderAll() {
    if (!state?.inventory?.length) {
      $("#emptyState").classList.remove("hidden");$("#workspace").classList.add("hidden");return;
    }
    $("#emptyState").classList.add("hidden");$("#workspace").classList.remove("hidden");
    fillFilters();renderDashboard();renderInventory();renderGeneral();renderWarehouse();renderNetworks();renderMaterials();renderSites();renderPeople();renderMovements();renderReports();wireTables();
  }

  function showModal(title,subtitle,html) {
    $("#modalTitle").textContent=title;$("#modalSubtitle").textContent=subtitle||"";$("#modalBody").innerHTML=html;
    $("#modal").classList.add("open");$("#modal").setAttribute("aria-hidden","false");
  }
  function closeModal(){ $("#modal").classList.remove("open");$("#modal").setAttribute("aria-hidden","true") }

  function openAsset(id) {
    const r=activeInventory().find(x=>x.id===id); if(!r)return; selectedId=id;
    const fields=[
      ["Código TIC",r.codigo||"SIN CÓDIGO"],["Código padre",r.codigoPadre],["Equipo / Material",r.equipo],["Descripción",r.descripcion],
      ["Marca",r.marca],["Modelo",r.modelo],["Serie o código",r.serie],["Sede",r.sede],["Área / Aula",r.area],
      ["Responsable",r.responsable],["DNI",r.dni],["Estado",r.estado],["Condición",r.condicion],["Situación",r.situacion],
      ["Origen Excel",r.source],["Observaciones",r.observaciones]
    ];
    const history=(r.history||[]).map(h=>`<div class="history-item"><b>${esc(h.source||h.type||"Movimiento")}</b> · ${esc([h.sede,h.area].filter(Boolean).join(" / "))}${h.responsable?` · ${esc(h.responsable)}`:""}<br><small>${esc(h.fecha||"")} ${esc(h.observaciones||"")}</small></div>`).join("");
    const duplicateNotice=r.duplicateSources>1?`<div class="alert info"><b>Código consolidado:</b> este código aparecía ${r.duplicateSources} veces en las hojas activas del Excel. El sistema lo mantiene como un solo registro.</div>`:"";
    showModal(`${r.codigo||r.id} · ${r.equipo||"Equipo"}`,`${r.sede||""} ${r.area?"/ "+r.area:""}`,`
      ${duplicateNotice}
      <div class="detail-grid">${fields.map(([k,v])=>`<div class="detail-field"><span>${esc(k)}</span><b>${esc(v||"—")}</b></div>`).join("")}</div>
      <div class="modal-actions">
        <button id="moveAsset" class="btn btn-primary">Mover / transferir</button>
        <button id="warehouseAsset" class="btn btn-warning">Enviar al almacén</button>
      </div>
      <h4 class="section-title">Historial del Código TIC</h4>${history||`<div class="muted">Sin historial adicional.</div>`}
    `);
    $("#moveAsset").onclick=()=>openMove(r,false);
    $("#warehouseAsset").onclick=()=>openMove(r,true);
  }

  function openMove(r,toWarehouse=false) {
    showModal(`Mover ${r.codigo||r.id}`,`Ubicación actual: ${[r.sede,r.area].filter(Boolean).join(" / ")}`,`
      <div id="formAlert"></div>
      <div class="form-grid">
        <div class="form-field"><label>Nueva sede</label><input id="mSede" value="${esc(r.sede||"")}"></div>
        <div class="form-field"><label>Nueva área / aula</label><input id="mArea" value="${esc(toWarehouse?"TIC / ALMACÉN":r.area||"")}"></div>
        <div class="form-field full"><label>Nuevo responsable</label><input id="mResp" value="${esc(toWarehouse?"":r.responsable||"")}" placeholder="Dejar vacío si vuelve a almacén"></div>
        <div class="form-field full"><label>Motivo / observación</label><textarea id="mObs" placeholder="Ej. devolución, cambio de responsable, traslado de sede..."></textarea></div>
      </div>
      <div class="modal-actions"><button id="saveMove" class="btn btn-primary">Confirmar movimiento</button></div>`);
    $("#saveMove").onclick=async()=>{
      const oldSite=r.sede, oldArea=r.area, oldResp=r.responsable;
      const newSite=$("#mSede").value.trim().toUpperCase(), newArea=$("#mArea").value.trim(), newResp=$("#mResp").value.trim(), motivo=$("#mObs").value.trim();
      r.history=r.history||[];
      r.history.unshift({type:"WEB",source:"Movimiento web",fecha:today(),sede:oldSite,area:oldArea,responsable:oldResp,estado:r.estado,observaciones:`Ubicación anterior. ${motivo}`});
      r.sede=newSite;r.area=newArea;r.responsable=newResp;r.locationType=norm(newArea).includes("ALMAC")?"ALMACEN":newResp?"ASIGNADO":"SEDE";
      state.webMovements=state.webMovements||[];
      state.webMovements.unshift({id:crypto.randomUUID(),fecha:today(),source:"Movimiento web",codigo:r.codigo,equipo:r.equipo,from:[oldSite,oldArea,oldResp].filter(Boolean).join(" / "),to:[newSite,newArea,newResp].filter(Boolean).join(" / "),responsable:newResp,motivo,observaciones:motivo});
      await persist();closeModal();renderAll();toast("Movimiento guardado sin duplicar el Código TIC.","success");
    };
  }

  function openPerson(key) {
    const p=collaborators().find(x=>x.key===key);if(!p)return;
    const items=p.items.sort((a,b)=>(a.codigo||a.id).localeCompare(b.codigo||b.id));
    showModal(p.name,`${items.length} equipos / registros a cargo`,`
      <div class="detail-grid">
        <div class="detail-field"><span>DNI</span><b>${esc(p.dni||"—")}</b></div>
        <div class="detail-field"><span>Sede(s)</span><b>${esc([...p.sites].join(", ")||"—")}</b></div>
        <div class="detail-field"><span>Área(s)</span><b>${esc([...p.areas].join(", ")||"—")}</b></div>
      </div>
      <div class="modal-actions"><button id="printCargo" class="btn btn-primary">Imprimir acta de cargo</button><button id="filterPerson" class="btn btn-soft">Ver en inventario</button></div>
      <h4 class="section-title">Equipos a cargo</h4>${tableHtml(items,false)}
    `);
    $("#printCargo").onclick=()=>printCargo(p);
    $("#filterPerson").onclick=()=>{closeModal();setView("inventory");$("#inventorySearch").value=p.name;page=1;renderInventory()};
    wireTables();
  }

  function printCargo(p) {
    const rows=p.items.map((r,i)=>`<tr><td>${i+1}</td><td>${esc(r.codigo||"SIN CÓDIGO")}</td><td>${esc(r.equipo)}</td><td>${esc(r.marca)}</td><td>${esc(r.modelo)}</td><td>${esc(r.serie)}</td><td>${esc(r.estado)}</td></tr>`).join("");
    const w=window.open("","_blank");
    w.document.write(`<!doctype html><html><head><title>Acta de Cargo</title><style>body{font-family:Arial;padding:30px;color:#10243d}.head{display:flex;align-items:center;gap:20px;border-bottom:3px solid #f7961f;padding-bottom:14px}.head img{width:170px}h1{font-size:20px}table{border-collapse:collapse;width:100%;font-size:11px;margin-top:18px}th,td{border:1px solid #bbb;padding:7px;text-align:left}.sign{display:grid;grid-template-columns:1fr 1fr;gap:80px;margin-top:70px;text-align:center}.line{border-top:1px solid #333;padding-top:8px}</style></head><body><div class="head"><img src="${document.querySelector(".brand img").src}"><div><h1>ACTA DE CARGO DE EQUIPOS TIC</h1><b>Colegio Praxis</b></div></div><p><b>Responsable:</b> ${esc(p.name)}<br><b>DNI:</b> ${esc(p.dni||"—")}<br><b>Sede:</b> ${esc([...p.sites].join(", "))}<br><b>Fecha:</b> ${today()}</p><table><thead><tr><th>N°</th><th>Código TIC</th><th>Equipo</th><th>Marca</th><th>Modelo</th><th>Serie</th><th>Estado</th></tr></thead><tbody>${rows}</tbody></table><div class="sign"><div class="line">Firma del colaborador</div><div class="line">Área TIC</div></div><script>setTimeout(()=>window.print(),500)<\/script></body></html>`);
    w.document.close();
  }

  function openAddAsset() {
    showModal("Agregar equipo","El Código TIC no puede repetirse.",`
      <div id="formAlert"></div>
      <div class="form-grid">
        <div class="form-field"><label>Código TIC *</label><input id="aCode" placeholder="Ej. MON-000100"></div>
        <div class="form-field"><label>Equipo / material *</label><input id="aEquipment"></div>
        <div class="form-field"><label>Sede</label><input id="aSite"></div>
        <div class="form-field"><label>Área / aula</label><input id="aArea"></div>
        <div class="form-field full"><label>Responsable</label><input id="aPerson"></div>
        <div class="form-field"><label>Marca</label><input id="aBrand"></div>
        <div class="form-field"><label>Modelo</label><input id="aModel"></div>
        <div class="form-field"><label>Serie</label><input id="aSerial"></div>
        <div class="form-field"><label>Estado</label><select id="aStatus"><option>OPERATIVO</option><option>EN MANTENIMIENTO</option><option>AVERIADO</option><option>BAJA</option></select></div>
        <div class="form-field full"><label>Observaciones</label><textarea id="aObs"></textarea></div>
      </div>
      <div class="modal-actions"><button id="saveAsset" class="btn btn-primary">Guardar equipo</button></div>`);
    $("#saveAsset").onclick=async()=>{
      const code=PraxisExcel.cleanCode($("#aCode").value), equipment=norm($("#aEquipment").value);
      const alert=$("#formAlert");
      if(!code||!equipment){alert.className="alert error";alert.textContent="Código TIC y Equipo / Material son obligatorios.";return}
      const existing=activeInventory().find(r=>norm(r.codigo)===code);
      if(existing){alert.className="alert error";alert.innerHTML=`El Código TIC <b>${esc(code)}</b> ya existe en ${esc(existing.sede)} / ${esc(existing.area)}. No se puede duplicar.`;return}
      const serial=$("#aSerial").value.trim();
      const duplicateSerial=serial&&activeInventory().find(r=>r.serie&&norm(r.serie)===norm(serial));
      if(duplicateSerial&&!confirm(`La serie ${serial} ya aparece en ${duplicateSerial.codigo||duplicateSerial.id}. ¿Deseas continuar?`))return;
      const r={id:code,codigo:code,codigoPadre:"",equipo:equipment,descripcion:"",marca:norm($("#aBrand").value),modelo:$("#aModel").value.trim(),serie:serial,sede:norm($("#aSite").value),area:$("#aArea").value.trim(),responsable:$("#aPerson").value.trim(),dni:"",estado:$("#aStatus").value,condicion:"",situacion:"INTERNO",source:"Registro web",observaciones:$("#aObs").value.trim(),duplicateSources:1,locationType:$("#aPerson").value.trim()?"ASIGNADO":norm($("#aArea").value).includes("ALMAC")?"ALMACEN":"SEDE",history:[{type:"WEB",source:"Registro web",fecha:today(),sede:norm($("#aSite").value),area:$("#aArea").value.trim(),responsable:$("#aPerson").value.trim(),estado:$("#aStatus").value,observaciones:"Registro creado desde la web"}]};
      state.inventory.unshift(r);state.webMovements.unshift({id:crypto.randomUUID(),fecha:today(),source:"Alta web",codigo:code,equipo:equipment,to:[r.sede,r.area,r.responsable].filter(Boolean).join(" / "),observaciones:r.observaciones});
      await persist();closeModal();renderAll();toast("Equipo registrado.","success");
    };
  }

  function csvDownload(records, filename, columns) {
    const lines=[columns.map(c=>`"${c[0]}"`).join(";")];
    records.forEach(r=>lines.push(columns.map(([_,key])=>`"${String(r[key]??"").replace(/"/g,'""')}"`).join(";")));
    const blob=new Blob(["\ufeff"+lines.join("\n")],{type:"text/csv;charset=utf-8"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=filename;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);
  }

  function exportInventory() {
    csvDownload(inventoryFiltered(),"Inventario_TIC_Praxis.csv",[["Código TIC","codigo"],["Código Padre","codigoPadre"],["Equipo","equipo"],["Marca","marca"],["Modelo","modelo"],["Serie","serie"],["Sede","sede"],["Área","area"],["Responsable","responsable"],["DNI","dni"],["Estado","estado"],["Condición","condicion"],["Situación","situacion"],["Observaciones","observaciones"],["Origen","source"]]);
  }

  function exportMovements() {
    csvDownload([...(state.webMovements||[]),...(state.transactions||[])],"Movimientos_TIC_Praxis.csv",[["Fecha","fecha"],["Origen","source"],["Código TIC","codigo"],["Equipo","equipo"],["Desde","from"],["Hacia","to"],["Responsable","responsable"],["Observaciones","observaciones"]]);
  }

  function downloadBackup() {
    const blob=new Blob([JSON.stringify(state,null,2)],{type:"application/json"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`Backup_Inventario_Praxis_${today()}.json`;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);
  }

  async function importExcel(file) {
    try{
      toast("Procesando Excel…");
      const next=await PraxisExcel.fromFile(file);
      if (!next.inventory.length) throw new Error("No se encontraron registros de inventario.");
      state=next;await persist();page=1;renderAll();setView("dashboard");
      toast(`Inventario importado: ${state.inventory.length} registros maestros.`,"success");
    }catch(err){console.error(err);toast(err.message||"No se pudo importar el Excel.","error")}
  }

  function demoInfo() {
    showModal("Cómo funciona","Diseñado para tu flujo de trabajo de Colegio Praxis.",`
      <div class="alert info">El sistema procesa las hojas <b>General, TIC - Almacén, Cámaras y Redes, Insum, mat y herra, Compras, Ventas y Préstamos</b>.</div>
      <div class="detail-grid">
        <div class="detail-field"><span>Código TIC</span><b>Único, no se duplica</b></div>
        <div class="detail-field"><span>Colaboradores</span><b>Equipos a cargo por persona</b></div>
        <div class="detail-field"><span>Almacén</span><b>Traslado sin crear otro código</b></div>
        <div class="detail-field"><span>Sedes</span><b>Filtro y resumen independiente</b></div>
        <div class="detail-field"><span>Historial</span><b>Conserva apariciones del Excel</b></div>
        <div class="detail-field"><span>Privacidad</span><b>Excel no se sube a GitHub</b></div>
      </div>`);
  }

  async function init() {
    state=await PraxisDB.get(STATE_KEY);
    renderAll();
    $$(".nav-item").forEach(b=>b.onclick=()=>setView(b.dataset.view));
    $$("[data-go]").forEach(b=>b.onclick=()=>setView(b.dataset.go));
    $("#mobileMenu").onclick=()=>$("#sidebar").classList.toggle("open");
    $("#modalClose").onclick=closeModal;$("#modal").addEventListener("click",e=>{if(e.target.id==="modal")closeModal()});
    $("#dismissPrivacy").onclick=()=>$("#privacyBanner").remove();
    $("#importBtnTop").onclick=()=>$("#excelInput").click();$("#importBtnEmpty").onclick=()=>$("#excelInput").click();
    $("#excelInput").onchange=e=>{const f=e.target.files[0];if(f)importExcel(f);e.target.value=""};
    $("#demoInfoBtn").onclick=demoInfo;
    $("#addAssetBtn").onclick=openAddAsset;$("#exportCsvBtn").onclick=exportInventory;$("#backupBtn").onclick=downloadBackup;$("#exportMovementsBtn").onclick=exportMovements;
    ["#inventorySearch","#siteFilter","#categoryFilter","#statusFilter","#locationFilter"].forEach(sel=>$(sel).addEventListener("input",()=>{page=1;renderInventory()}));
    $("#peopleSearch").addEventListener("input",renderPeople);
    $("#globalSearch").addEventListener("input",e=>{
      if(!state?.inventory?.length)return;
      $("#inventorySearch").value=e.target.value;page=1;if(e.target.value.trim())setView("inventory");renderInventory();
    });
  }

  document.addEventListener("DOMContentLoaded",init);
})();
