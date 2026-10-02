
(() => {
  const STATE_KEY = "praxis-state-v2";
  const PAGE_SIZE = 30;
  const OWNER_ADMIN_EMAIL = "r.cardenas@praxis.edu.pe";
  let state = null;
  let currentView = "dashboard";
  let page = 1;
  let selectedId = null;
  let cloudMode = false;
  let currentProfile = null;
  let realtimeChannel = null;
  let accessChannel = null;

  const $ = s => document.querySelector(s);
  const $$ = s => [...document.querySelectorAll(s)];
  const esc = value => String(value ?? "").replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
  const norm = value => PraxisExcel.norm(value);
  const today = () => {
    const d=new Date();
    const y=d.getFullYear(), m=String(d.getMonth()+1).padStart(2,"0"), day=String(d.getDate()).padStart(2,"0");
    return `${y}-${m}-${day}`;
  };

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

  async function persist(action="ACTUALIZAR_INVENTARIO") {
    if (!cloudMode || !PraxisCloud.configured()) {
      await PraxisDB.set(STATE_KEY,state);
      return true;
    }

    try {
      await PraxisCloud.saveState(state,action);
      try{ await PraxisDB.set(STATE_KEY,state); }catch(err){ console.warn("No se pudo actualizar la caché local",err); }
      return true;
    } catch (err) {
      console.error(err);
      try{
        const latest=await PraxisCloud.loadState();
        state=latest.state;
        try{ if(state) await PraxisDB.set(STATE_KEY,state); }catch(_){}
        renderAll();
      }catch(reloadErr){
        console.error("No se pudo restaurar el estado remoto",reloadErr);
      }

      const raw=String(err?.message||"");
      if(raw==="VERSION_CONFLICT" || raw.includes("VERSION_CONFLICT")){
        toast("Otro usuario actualizó el inventario. Se recargó la versión más reciente.","error");
      }else if(raw.includes("READ_ONLY_ROLE")){
        toast("Tu usuario es de Solo lectura. No puede realizar cambios.","error");
      }else{
        toast(raw||"No se pudo guardar en la base de datos.","error");
      }
      throw err;
    }
  }

  function canEdit() {
    return !cloudMode || PraxisCloud.canEdit();
  }

  function isAdmin() {
    return cloudMode && PraxisCloud.isAdmin();
  }

  function canManageUsers() {
    const p=currentProfile || PraxisCloud.getProfile();
    return Boolean(cloudMode && p?.activo && String(p.email||"").toLowerCase()===OWNER_ADMIN_EMAIL);
  }

  function ensureEditable() {
    if (canEdit()) return true;
    toast("Tu usuario es de Solo lectura. No puede modificar el inventario.","error");
    return false;
  }

  const roleLabel = role => ({
    ADMIN_TIC:"Administrador TIC",
    EDITOR:"Editor",
    CONSULTA:"Solo lectura"
  }[role] || role || "Local");

  const fmtDateTime = value => {
    if(!value) return "Nunca";
    const d=new Date(value);
    return Number.isNaN(d.getTime()) ? String(value) : d.toLocaleString("es-PE",{dateStyle:"short",timeStyle:"short"});
  };

  function updateProfileUI() {
    const p = currentProfile || PraxisCloud.getProfile();
    const displayName=(p && (p.nombre || p.email)) || "Usuario";
    const displayRole=(p && String(p.email||"").toLowerCase()===OWNER_ADMIN_EMAIL) ? "Master TIC" : roleLabel((p && p.role) || "LOCAL");
    if ($("#profileName")) $("#profileName").textContent = displayName;
    if ($("#profileRole")) $("#profileRole").textContent = displayRole;
    if ($("#sidebarProfileName")) $("#sidebarProfileName").textContent = displayName;
    if ($("#sidebarProfileRole")) $("#sidebarProfileRole").textContent = displayRole;
    $$(".edit-only").forEach(el=>el.classList.toggle("hidden",cloudMode && !canEdit()));
    $$(".admin-only").forEach(el=>el.classList.toggle("hidden",!canManageUsers()));
    $$(".cloud-only").forEach(el=>el.classList.toggle("hidden",!cloudMode));
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
      people:["Colaboradores","Alta, asignación y transferencia de equipos por responsable"],
      warehouse:["Almacén TIC","Control de equipos almacenados"],
      networks:["Cámaras y redes","Infraestructura tecnológica y seguridad"],
      materials:["Materiales","Insumos, materiales y herramientas"],
      movements:["Movimientos","Transferencias e historial"],
      review:["Revisión","Conflictos, códigos faltantes y depuración"],
      reports:["Reportes","Resumen del inventario"],
      users:["Usuarios y roles","Administración de accesos y auditoría"]
    };
    $("#pageTitle").textContent = titles[view][0];
    $("#pageSubtitle").textContent = titles[view][1];
    if($("#globalSearch")){
      $("#globalSearch").placeholder = view==="people"
        ? "Buscar colaborador por nombre o DNI..."
        : "Buscar código TIC, equipo, serie, colaborador, sede, área...";
    }
    if (view === "users") { if(!canManageUsers()){ setView("dashboard"); return; } renderUsers(); renderCargoAudit(); renderAudit(); }
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
      [inv.filter(r=>r.needsReview).length,"Ubicación por revisar"],
      [inv.filter(r=>r.serie).length,"Con serie/código"],
      [state.transactions?.length||0,"Registros históricos"],
      [state.webMovements?.length||0,"Movimientos web"]
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
        <td><span class="badge ${statusClass(r.estado)}">${esc(r.estado||"SIN ESTADO")}</span><div class="muted">${locationLabel(r)}</div>${r.needsReview?'<div class="muted" style="color:#b15c00;font-weight:800">⚠ Revisar ubicación</div>':""}</td>
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
    const rows=activeInventory().filter(r=>r.source==="General");
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

  function reviewItems() {
    return activeInventory().filter(r => r.needsReview || !r.codigo);
  }

  function renderReview() {
    const q = norm($("#reviewSearch")?.value || "");
    const all = reviewItems();
    const items = all.filter(r => {
      if (!q) return true;
      return norm([r.codigo,r.id,r.equipo,r.marca,r.modelo,r.serie,r.sede,r.area,r.responsable,r.dni].join(" ")).includes(q);
    });
    const conflicts = activeInventory().filter(r=>r.needsReview).length;
    const noCode = activeInventory().filter(r=>!r.codigo).length;
    const consolidated = activeInventory().filter(r=>(r.duplicateSources||0)>1).length;
    $("#reviewSummary").innerHTML = [
      [all.length,"Pendientes de revisión"],
      [conflicts,"Ubicaciones conflictivas"],
      [noCode,"Sin Código TIC"],
      [consolidated,"Códigos consolidados"]
    ].map(([n,l])=>`<div class="quality-box"><b>${n}</b><span>${esc(l)}</span></div>`).join("");

    const rows = items.map(r => {
      const issues = [
        r.needsReview ? '<span class="badge warn">UBICACIÓN</span>' : '',
        !r.codigo ? '<span class="badge bad">SIN CÓDIGO TIC</span>' : ''
      ].filter(Boolean).join(" ");
      return `<tr>
        <td><span class="code">${esc(r.codigo||r.id)}</span></td>
        <td>${issues}</td>
        <td><b>${esc(r.equipo||"Sin tipo")}</b><div class="muted">${esc([r.marca,r.modelo].filter(Boolean).join(" / "))}</div></td>
        <td>${esc(r.sede)}</td>
        <td>${esc(r.area)}</td>
        <td>${esc(r.responsable)}</td>
        <td>${r.duplicateSources>1?`${r.duplicateSources} apariciones`:"1 aparición"}</td>
        <td><button class="btn btn-soft" data-review="${esc(r.id)}">Revisar</button></td>
      </tr>`;
    }).join("");
    $("#reviewTable").innerHTML = `<div class="table-wrap"><table><thead><tr><th>Código</th><th>Observación</th><th>Equipo</th><th>Sede</th><th>Área</th><th>Responsable</th><th>Origen</th><th></th></tr></thead><tbody>${rows||'<tr><td colspan="8">No hay registros pendientes con este filtro.</td></tr>'}</tbody></table></div>`;
    $$("[data-review]").forEach(b=>b.onclick=()=>openAsset(b.dataset.review));
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
    fillFilters();renderDashboard();renderInventory();renderGeneral();renderWarehouse();renderNetworks();renderMaterials();renderSites();renderPeople();renderMovements();renderReview();renderReports();wireTables();
  }

  function showModal(title,subtitle,html) {
    $("#modalTitle").textContent=title;$("#modalSubtitle").textContent=subtitle||"";$("#modalBody").innerHTML=html;
    $("#modal").classList.add("open");$("#modal").setAttribute("aria-hidden","false");
  }
  function closeModal(){ $("#modal").classList.remove("open");$("#modal").setAttribute("aria-hidden","true") }

  function openAddAsset() {
    if(!ensureEditable()) return;
    showModal("Agregar equipo","Crea un único registro maestro. El Código TIC no puede repetirse.",`
      <div id="formAlert"></div>
      <div class="form-grid">
        <div class="form-field"><label>Código TIC *</label><input id="nCode" placeholder="Ej. MON-000100"></div>
        <div class="form-field"><label>Código Padre TIC</label><input id="nParent"></div>
        <div class="form-field"><label>Equipo / material *</label><input id="nEquipment" placeholder="Ej. MONITOR"></div>
        <div class="form-field"><label>Descripción</label><input id="nDescription"></div>
        <div class="form-field"><label>Marca</label><input id="nBrand"></div>
        <div class="form-field"><label>Modelo</label><input id="nModel"></div>
        <div class="form-field"><label>Serie o código</label><input id="nSerial"></div>
        <div class="form-field"><label>Estado</label><input id="nStatus" value="OPERATIVO"></div>
        <div class="form-field"><label>Condición</label><input id="nCondition"></div>
        <div class="form-field"><label>Situación</label><input id="nSituation" value="INTERNO"></div>
        <div class="form-field"><label>Sede</label><input id="nSite"></div>
        <div class="form-field"><label>Área / aula</label><input id="nArea"></div>
        <div class="form-field full"><label>Responsable</label><input id="nPerson"></div>
        <div class="form-field"><label>DNI</label><input id="nDni" maxlength="8"></div>
        <div class="form-field full"><label>Observaciones</label><textarea id="nObs"></textarea></div>
      </div>
      <div class="modal-actions"><button id="saveNewAsset" class="btn btn-primary">Guardar equipo</button></div>
    `);

    $("#saveNewAsset").onclick=async()=>{
      const alert=$("#formAlert");
      const code=PraxisExcel.cleanCode($("#nCode").value);
      const equipment=norm($("#nEquipment").value);
      if(!code){alert.className="alert error";alert.textContent="El Código TIC es obligatorio.";return}
      if(!equipment){alert.className="alert error";alert.textContent="Equipo / Material es obligatorio.";return}
      const duplicate=activeInventory().find(x=>x.codigo && norm(x.codigo)===code);
      if(duplicate){
        alert.className="alert error";
        alert.innerHTML=`El Código TIC <b>${esc(code)}</b> ya existe. Abre ese registro y actualízalo; no se creará un duplicado.`;
        return;
      }
      const serial=$("#nSerial").value.trim();
      const duplicateSerial=serial&&activeInventory().find(x=>x.serie&&norm(x.serie)===norm(serial));
      if(duplicateSerial&&!confirm(`La serie ${serial} ya aparece en ${duplicateSerial.codigo||duplicateSerial.id}. ¿Deseas continuar?`)) return;

      const responsable=$("#nPerson").value.trim();
      const area=$("#nArea").value.trim();
      const record={
        id:code,codigo:code,codigoPadre:PraxisExcel.cleanCode($("#nParent").value),
        equipo:equipment,descripcion:$("#nDescription").value.trim(),detalle:"",
        marca:norm($("#nBrand").value),modelo:$("#nModel").value.trim(),serie:serial,
        sede:norm($("#nSite").value),area,responsable,dni:$("#nDni").value.trim(),
        estado:norm($("#nStatus").value),condicion:norm($("#nCondition").value),
        situacion:norm($("#nSituation").value),observaciones:$("#nObs").value.trim(),
        source:"Web",cantidad:"1",
        locationType:norm(area).includes("ALMAC")?"ALMACEN":responsable?"ASIGNADO":"SEDE",
        needsReview:false,conflictLocations:[],conflictRecords:[],duplicateSources:1,
        history:[{type:"WEB",source:"Alta web",fecha:today(),sede:norm($("#nSite").value),area,responsable,estado:norm($("#nStatus").value),observaciones:"Registro creado desde Inventario Praxis"}]
      };
      state.inventory=state.inventory||[];
      state.inventory.push(record);
      state.webMovements=state.webMovements||[];
      state.webMovements.unshift({id:crypto.randomUUID(),fecha:today(),source:"Alta web",codigo:code,equipo:equipment,from:"",to:[record.sede,area,responsable].filter(Boolean).join(" / "),responsable,observaciones:record.observaciones});
      await persist("ALTA_EQUIPO");
      closeModal();page=1;renderAll();toast("Equipo agregado sin duplicar el Código TIC.","success");
    };
  }

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
    const conflictNotice=r.needsReview?`<div class="alert error"><b>Revisar ubicación actual:</b> el mismo Código TIC aparece asociado a más de una sede, área o responsable en el Excel. Verifica el historial antes de confirmar una transferencia.<br><span class="muted">${esc((r.conflictLocations||[]).join(" · "))}</span></div>`:"";
    const actionButtons = canEdit() ? `
      <div class="modal-actions">
        <button id="editAsset" class="btn btn-soft">Editar ficha</button>
        <button id="moveAsset" class="btn btn-primary">Mover / transferir</button>
        <button id="warehouseAsset" class="btn btn-warning">Enviar al almacén</button>
        ${r.needsReview?'<button id="resolveAsset" class="btn btn-danger">Resolver ubicación</button>':""}
      </div>` : `<div class="readonly-banner">Modo CONSULTA: puedes ver la ficha y el historial, pero no modificar datos.</div>`;
    showModal(`${r.codigo||r.id} · ${r.equipo||"Equipo"}`,`${r.sede||""} ${r.area?"/ "+r.area:""}`,`
      ${duplicateNotice}
      ${conflictNotice}
      <div class="detail-grid">${fields.map(([k,v])=>`<div class="detail-field"><span>${esc(k)}</span><b>${esc(v||"—")}</b></div>`).join("")}</div>
      ${actionButtons}
      <h4 class="section-title">Historial del Código TIC</h4>${history||`<div class="muted">Sin historial adicional.</div>`}
    `);
    if (canEdit()) {
      $("#editAsset").onclick=()=>openEditAsset(r);
      $("#moveAsset").onclick=()=>openMove(r,false);
      $("#warehouseAsset").onclick=()=>openMove(r,true);
      if (r.needsReview) $("#resolveAsset").onclick=()=>openResolveConflict(r);
    }
  }

  function openEditAsset(r) {
    if(!ensureEditable()) return;
    showModal(`Editar ${r.codigo||r.id}`,"Actualiza la ficha sin crear un segundo registro.",`
      <div id="formAlert"></div>
      <div class="form-grid">
        <div class="form-field"><label>Código TIC</label><input id="eCode" value="${esc(r.codigo||"")}" placeholder="Ej. MON-000100"></div>
        <div class="form-field"><label>Código Padre TIC</label><input id="eParent" value="${esc(r.codigoPadre||"")}"></div>
        <div class="form-field"><label>Equipo / material *</label><input id="eEquipment" value="${esc(r.equipo||"")}"></div>
        <div class="form-field"><label>Descripción</label><input id="eDescription" value="${esc(r.descripcion||"")}"></div>
        <div class="form-field"><label>Marca</label><input id="eBrand" value="${esc(r.marca||"")}"></div>
        <div class="form-field"><label>Modelo</label><input id="eModel" value="${esc(r.modelo||"")}"></div>
        <div class="form-field"><label>Serie o código</label><input id="eSerial" value="${esc(r.serie||"")}"></div>
        <div class="form-field"><label>Estado</label><input id="eStatus" value="${esc(r.estado||"")}"></div>
        <div class="form-field"><label>Condición</label><input id="eCondition" value="${esc(r.condicion||"")}"></div>
        <div class="form-field"><label>Situación</label><input id="eSituation" value="${esc(r.situacion||"")}"></div>
        <div class="form-field"><label>Sede</label><input id="eSite" value="${esc(r.sede||"")}"></div>
        <div class="form-field"><label>Área / aula</label><input id="eArea" value="${esc(r.area||"")}"></div>
        <div class="form-field full"><label>Responsable</label><input id="ePerson" value="${esc(r.responsable||"")}"></div>
        <div class="form-field"><label>DNI</label><input id="eDni" value="${esc(r.dni||"")}"></div>
        <div class="form-field full"><label>Observaciones</label><textarea id="eObs">${esc(r.observaciones||"")}</textarea></div>
      </div>
      <div class="modal-actions"><button id="saveEditAsset" class="btn btn-primary">Guardar cambios</button></div>
    `);

    $("#saveEditAsset").onclick=async()=>{
      const alert=$("#formAlert");
      const newCode=PraxisExcel.cleanCode($("#eCode").value);
      const equipment=norm($("#eEquipment").value);
      if(!equipment){alert.className="alert error";alert.textContent="Equipo / Material es obligatorio.";return}
      if(r.codigo && !newCode){alert.className="alert error";alert.textContent="Un equipo que ya tiene Código TIC no puede quedar sin código.";return}
      if(newCode){
        const existing=activeInventory().find(x=>x!==r && norm(x.codigo)===newCode);
        if(existing){alert.className="alert error";alert.innerHTML=`El Código TIC <b>${esc(newCode)}</b> ya existe en ${esc(existing.sede)} / ${esc(existing.area)}. No se puede duplicar.`;return}
      }
      const serial=$("#eSerial").value.trim();
      const duplicateSerial=serial&&activeInventory().find(x=>x!==r&&x.serie&&norm(x.serie)===norm(serial));
      if(duplicateSerial&&!confirm(`La serie ${serial} ya aparece en ${duplicateSerial.codigo||duplicateSerial.id}. ¿Deseas continuar?`))return;

      const before=[r.sede,r.area,r.responsable].filter(Boolean).join(" / ");
      r.history=r.history||[];
      r.history.unshift({type:"WEB",source:"Edición web",fecha:today(),sede:r.sede,area:r.area,responsable:r.responsable,estado:r.estado,observaciones:"Ficha anterior antes de edición"});

      if(newCode && newCode!==r.codigo){ r.codigo=newCode; r.id=newCode; }
      r.codigoPadre=PraxisExcel.cleanCode($("#eParent").value);
      r.equipo=equipment;
      r.descripcion=$("#eDescription").value.trim();
      r.marca=norm($("#eBrand").value);
      r.modelo=$("#eModel").value.trim();
      r.serie=serial;
      r.estado=norm($("#eStatus").value);
      r.condicion=norm($("#eCondition").value);
      r.situacion=norm($("#eSituation").value);
      r.sede=norm($("#eSite").value);
      r.area=$("#eArea").value.trim();
      r.responsable=$("#ePerson").value.trim();
      r.dni=$("#eDni").value.trim();
      r.observaciones=$("#eObs").value.trim();
      r.locationType=norm(r.area).includes("ALMAC")?"ALMACEN":r.responsable?"ASIGNADO":"SEDE";

      state.webMovements=state.webMovements||[];
      state.webMovements.unshift({
        id:crypto.randomUUID(),fecha:today(),source:"Edición web",codigo:r.codigo,equipo:r.equipo,
        from:before,to:[r.sede,r.area,r.responsable].filter(Boolean).join(" / "),
        responsable:r.responsable,observaciones:"Actualización de ficha"
      });
      await persist("EDITAR_FICHA");closeModal();renderAll();toast("Ficha actualizada y sincronizada.","success");
    };
  }

  function openResolveConflict(r) {
    if(!ensureEditable()) return;
    const candidates=(r.conflictRecords&&r.conflictRecords.length?r.conflictRecords:(r.history||[]).filter(h=>h.sede||h.area||h.responsable))
      .filter((x,i,a)=>a.findIndex(y=>norm([y.sede,y.area,y.responsable].join("|"))===norm([x.sede,x.area,x.responsable].join("|")))===i);

    const rows=candidates.map((x,i)=>`<tr>
      <td>${esc(x.source||"Excel")}</td>
      <td>${esc(x.fecha||"")}</td>
      <td>${esc(x.sede||"")}</td>
      <td>${esc(x.area||"")}</td>
      <td>${esc(x.responsable||"")}</td>
      <td>${esc(x.estado||"")}</td>
      <td><button class="btn btn-soft" data-pick-conflict="${i}">Usar esta ubicación</button></td>
    </tr>`).join("");

    showModal(`Resolver ${r.codigo||r.id}`,"Elige la ubicación vigente según el Excel o define una nueva manualmente.",`
      <div class="alert info"><b>No se creará otro Código TIC.</b> Solo se confirmará la ubicación vigente y se conservarán las demás apariciones en el historial.</div>
      <div class="table-wrap"><table><thead><tr><th>Hoja</th><th>Fecha</th><th>Sede</th><th>Área</th><th>Responsable</th><th>Estado</th><th></th></tr></thead><tbody>${rows||'<tr><td colspan="7">No hay ubicaciones estructuradas. Define la ubicación manualmente.</td></tr>'}</tbody></table></div>
      <div class="modal-actions"><button id="manualConflict" class="btn btn-primary">Definir ubicación manualmente</button></div>
    `);

    $$("[data-pick-conflict]").forEach(b=>b.onclick=async()=>{
      const x=candidates[Number(b.dataset.pickConflict)];
      const before=[r.sede,r.area,r.responsable].filter(Boolean).join(" / ");
      r.history=r.history||[];
      r.history.unshift({type:"WEB",source:"Resolución de conflicto",fecha:today(),sede:r.sede,area:r.area,responsable:r.responsable,estado:r.estado,observaciones:"Ubicación anterior antes de resolver conflicto"});
      r.sede=norm(x.sede||"");r.area=x.area||"";r.responsable=x.responsable||"";r.dni=x.dni||r.dni||"";
      if(x.estado) r.estado=norm(x.estado);
      if(x.condicion) r.condicion=norm(x.condicion);
      if(x.situacion) r.situacion=norm(x.situacion);
      if(x.source) r.source=x.source;
      r.locationType=norm(r.area).includes("ALMAC")?"ALMACEN":r.responsable?"ASIGNADO":"SEDE";
      r.needsReview=false;r.conflictLocations=[];r.conflictRecords=[];
      state.webMovements=state.webMovements||[];
      state.webMovements.unshift({id:crypto.randomUUID(),fecha:today(),source:"Resolución de conflicto",codigo:r.codigo,equipo:r.equipo,from:before,to:[r.sede,r.area,r.responsable].filter(Boolean).join(" / "),responsable:r.responsable,observaciones:`Ubicación confirmada desde ${x.source||"Excel"} fila ${x.row||""}`});
      await persist("RESOLVER_CONFLICTO");closeModal();renderAll();toast("Ubicación confirmada y sincronizada.","success");
    });
    $("#manualConflict").onclick=()=>openMove(r,false);
  }

  function openMove(r,toWarehouse=false) {
    if(!ensureEditable()) return;
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
      r.sede=newSite;r.area=newArea;r.responsable=newResp;
      if(toWarehouse || norm(newArea).includes("ALMAC")) r.dni="";
      r.locationType=norm(newArea).includes("ALMAC")?"ALMACEN":newResp?"ASIGNADO":"SEDE";
      r.needsReview=false;
      r.conflictLocations=[];
      r.conflictRecords=[];
      state.webMovements=state.webMovements||[];
      state.webMovements.unshift({id:crypto.randomUUID(),fecha:today(),source:"Movimiento web",codigo:r.codigo,equipo:r.equipo,from:[oldSite,oldArea,oldResp].filter(Boolean).join(" / "),to:[newSite,newArea,newResp].filter(Boolean).join(" / "),responsable:newResp,motivo,observaciones:motivo});
      await persist("MOVIMIENTO_EQUIPO");closeModal();renderAll();toast("Movimiento guardado y sincronizado sin duplicar el Código TIC.","success");
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
      <div class="modal-actions">
        <button id="editPersonCargo" class="btn btn-soft">Editar datos</button>
        <button id="printCargo" class="btn btn-primary">Imprimir cargo</button>
        <button id="filterPerson" class="btn btn-soft">Ver en inventario</button>
      </div>
      <h4 class="section-title">Equipos a cargo</h4>${tableHtml(items,false)}
    `);
    $("#editPersonCargo").onclick=()=>openEditPersonCargo(p);
    $("#printCargo").onclick=()=>printCargo(p);
    $("#filterPerson").onclick=()=>{closeModal();setView("inventory");$("#inventorySearch").value=p.name;page=1;renderInventory()};
    wireTables();
  }

  function cargoOverrideFor(p) {
    const key=cargoPersonKey(p);
    return (state?.cargoOverrides && state.cargoOverrides[key]) || {};
  }

  function cargoSoftwareDefaults(){
    return [
      {name:"Windows",version:"10",details:"22H2"},
      {name:"Ms Office Professional Plus",version:"2021",details:""},
      {name:"Winrar",version:"2021",details:""},
      {name:"VLC",version:"",details:""},
      {name:"Aimp3",version:"",details:""},
      {name:"Edge",version:"",details:""},
      {name:"Chrome",version:"",details:""},
      {name:"Firefox",version:"",details:""},
      {name:"Adobe Acrobat",version:"2021",details:""},
      {name:"Anydesk",version:"",details:""},
      {name:"",version:"",details:""},
      {name:"",version:"",details:""},
      {name:"",version:"",details:""},
      {name:"",version:"",details:""},
      {name:"",version:"",details:""}
    ];
  }
  
  function currentCargoDocForPerson(p){
    state.cargoDocuments=state.cargoDocuments||[];
    const key=cargoPersonKey(p);
    const base=cargoSnapshot(p.items);
    const fingerprint=cargoFingerprint(base);
    return [...state.cargoDocuments].reverse().find(d=>{
      const samePerson=d.personKey===key || norm(d.personName||'')===norm(p.name||'');
      return samePerson && (d.itemFingerprint===fingerprint || cargoFingerprint(d.items||[])===fingerprint);
    }) || null;
  }
  
  function latestPrintedCargoForPerson(p){
    const key=cargoPersonKey(p);
    const dni=String(p?.dni||"").replace(/\D/g,"");
    return [...(state?.cargoDocuments||[])].reverse().find(d=>{
      if(!d.printedAt) return false;
      const docDni=String(d.dni||"").replace(/\D/g,"");
      return d.personKey===key || (dni && docDni===dni);
    }) || null;
  }
  
  async function setCargoReceived(docId,received=true){
    if(!ensureEditable()) return;
    const doc=(state?.cargoDocuments||[]).find(d=>d.id===docId);
    if(!doc){toast("No se encontró el cargo.","error");return}
    doc.receivedSigned=Boolean(received);
    doc.receivedAt=received?new Date().toISOString():"";
    doc.receivedBy=currentProfile?.id||"";
    doc.receivedByName=received?(currentProfile?.nombre||currentProfile?.email||""):"";
    await persist("CONFIRMAR_RECEPCION_CARGO");
    toast(received?"Cargo marcado como firmado y equipos recibidos.":"Se quitó la confirmación de recepción.","success");
  }
  
  function openEditPersonCargo(p){
    if(!ensureEditable()) return;
  
    state.cargoOverrides=state.cargoOverrides||{};
    state.cargoDocuments=state.cargoDocuments||[];
  
    const key=cargoPersonKey(p);
    const saved=state.cargoOverrides[key]||{};
    const existingDoc=currentCargoDocForPerson(p);
    const baseItems=cargoSnapshot(p.items);
    const fingerprint=cargoFingerprint(baseItems);
    const docCode=existingDoc?.code || nextCargoCode();
  
    const sites=[...p.sites], areas=[...p.areas];
    const personName=saved.personName ?? existingDoc?.personName ?? p.name ?? "";
    const dni=saved.dni ?? existingDoc?.dni ?? p.dni ?? "";
    const sede=saved.sede ?? existingDoc?.sede ?? (sites.length===1?sites[0]:"");
    const area=saved.area ?? existingDoc?.area ?? (areas.length===1?areas[0]:"");
    const situacion=saved.situacion ?? existingDoc?.situacion ?? "INTERNO";
    const technician=saved.technicianName ?? existingDoc?.technicianName ?? "CARDENAS CURISINCHE, ROBERTO ALEJANDRO";
    const chief=saved.chiefName ?? existingDoc?.chiefName ?? TIC_HEAD_NAME;
    const initialDate=saved.initialDate ?? existingDoc?.initialDate ?? today();
    const finalDate=saved.finalDate ?? existingDoc?.finalDate ?? "";
    const initialObservation=saved.initialObservation ?? existingDoc?.initialObservation ?? "";
    const finalObservation=saved.finalObservation ?? existingDoc?.finalObservation ?? "";
  
    const savedItems=Array.isArray(saved.items)&&saved.items.length?saved.items:(existingDoc?.items||[]);
    const savedById=new Map(savedItems.map(x=>[x.id||x.codigo,x]));
    const items=baseItems.map(r=>({...r,...(savedById.get(r.id||r.codigo)||{})}));
  
    const softwareRaw=Array.isArray(saved.software)&&saved.software.length?saved.software:(existingDoc?.software||cargoSoftwareDefaults());
    const software=[...softwareRaw];
    while(software.length<15) software.push({name:"",version:"",details:""});
  
    showModal("Editar cargo técnico","Puedes editar todo el contenido de la ficha. El código FTEC es automático y no se puede modificar.",`
      <div class="alert info"><b>Código automático:</b> este cargo usará <b>${esc(docCode)}</b>. El siguiente continuará correlativamente.</div>
      <div id="formAlert"></div>
  
      <h4 class="section-title">Datos generales del cargo</h4>
      <div class="form-grid cargo-editor-general">
        <div class="form-field"><label>Código FTEC</label><input value="${esc(docCode)}" readonly class="readonly-code"></div>
        <div class="form-field"><label>Situación</label><input id="pcSituation" value="${esc(situacion)}"></div>
        <div class="form-field full"><label>Usuario / colaborador *</label><input id="pcName" value="${esc(personName)}"></div>
        <div class="form-field"><label>DNI</label><input id="pcDni" maxlength="8" value="${esc(dni)}"></div>
        <div class="form-field"><label>Sede</label><input id="pcSite" value="${esc(sede)}"></div>
        <div class="form-field full"><label>Área</label><input id="pcArea" value="${esc(area)}"></div>
        <div class="form-field"><label>Fecha revisión inicial</label><input id="pcInitialDate" type="date" value="${esc(initialDate)}"></div>
        <div class="form-field"><label>Fecha revisión final</label><input id="pcFinalDate" type="date" value="${esc(finalDate)}"></div>
        <div class="form-field full"><label>Observación revisión inicial</label><textarea id="pcInitialObs">${esc(initialObservation)}</textarea></div>
        <div class="form-field full"><label>Observación revisión final</label><textarea id="pcFinalObs">${esc(finalObservation)}</textarea></div>
        <div class="form-field full"><label>Técnico / Responsable de Inf. Tec.</label><input id="pcTech" value="${esc(technician)}"></div>
        <div class="form-field full"><label>Jefe de TIC</label><input id="pcChief" value="${esc(chief)}"></div>
      </div>
  
      <h4 class="section-title">Hardware del cargo</h4>
      <div class="table-wrap cargo-editor-table">
        <table>
          <thead><tr><th>#</th><th>Cant.</th><th>Equipo</th><th>Tipo</th><th>Marca</th><th>Modelo</th><th>Características</th><th>Código TIC</th><th>Serie</th><th>Estado</th><th>Observación</th></tr></thead>
          <tbody>${items.map((r,i)=>`<tr data-cargo-row="${i}">
            <td>${i+1}</td>
            <td><input data-cargo-field="cantidad" value="${esc(r.cantidad||"1")}"></td>
            <td><input data-cargo-field="equipo" value="${esc(r.equipo||"")}"></td>
            <td><input data-cargo-field="bien" value="${esc(r.bien||cargoItemType(r))}"></td>
            <td><input data-cargo-field="marca" value="${esc(r.marca||"")}"></td>
            <td><input data-cargo-field="modelo" value="${esc(r.modelo||"")}"></td>
            <td><input data-cargo-field="caracteristicas" value="${esc(r.caracteristicas||"")}"></td>
            <td><input data-cargo-field="codigo" value="${esc(r.codigo||"")}"></td>
            <td><input data-cargo-field="serie" value="${esc(r.serie||"")}"></td>
            <td><input data-cargo-field="estado" value="${esc(r.estado||"")}"></td>
            <td><input data-cargo-field="observaciones" value="${esc(r.observaciones||"")}"></td>
          </tr>`).join("")}</tbody>
        </table>
      </div>
  
      <h4 class="section-title">Software del cargo</h4>
      <div class="table-wrap cargo-software-table">
        <table>
          <thead><tr><th>#</th><th>Nombre</th><th>Versión</th><th>Detalles</th></tr></thead>
          <tbody>${software.slice(0,15).map((s,i)=>`<tr data-software-row="${i}">
            <td>${i+1}</td>
            <td><input data-software-field="name" value="${esc(s.name||"")}"></td>
            <td><input data-software-field="version" value="${esc(s.version||"")}"></td>
            <td><input data-software-field="details" value="${esc(s.details||"")}"></td>
          </tr>`).join("")}</tbody>
        </table>
      </div>
  
      <div class="modal-actions cargo-editor-actions">
        <button id="savePersonCargo" class="btn btn-primary">Guardar cambios del cargo</button>
      </div>
    `);
  
    $("#savePersonCargo").onclick=async()=>{
      const alert=$("#formAlert");
      const name=$("#pcName").value.trim();
      const cargoDni=$("#pcDni").value.replace(/\D/g,"").slice(0,8);
      const cargoSede=$("#pcSite").value.trim();
      const cargoArea=$("#pcArea").value.trim();
      const tech=$("#pcTech").value.trim();
      const chiefName=$("#pcChief").value.trim();
  
      if(!name){alert.className="alert error";alert.textContent="El nombre del usuario es obligatorio.";return}
      if(cargoDni && !/^\d{8}$/.test(cargoDni)){alert.className="alert error";alert.textContent="El DNI debe tener 8 dígitos.";return}
      if(!tech || !chiefName){alert.className="alert error";alert.textContent="Completa los nombres del técnico y del Jefe de TIC.";return}
  
      const editedItems=items.map((base,i)=>{
        const row=document.querySelector(`[data-cargo-row="${i}"]`);
        const get=f=>row?.querySelector(`[data-cargo-field="${f}"]`)?.value.trim()||"";
        return {
          ...base,
          cantidad:get("cantidad")||"1",
          equipo:get("equipo"),
          bien:get("bien"),
          marca:get("marca"),
          modelo:get("modelo"),
          caracteristicas:get("caracteristicas"),
          codigo:get("codigo"),
          serie:get("serie"),
          estado:get("estado"),
          observaciones:get("observaciones")
        };
      });
  
      const editedSoftware=software.slice(0,15).map((base,i)=>{
        const row=document.querySelector(`[data-software-row="${i}"]`);
        const get=f=>row?.querySelector(`[data-software-field="${f}"]`)?.value.trim()||"";
        return {name:get("name"),version:get("version"),details:get("details")};
      });
  
      const payload={
        personName:name,
        dni:cargoDni,
        sede:cargoSede,
        area:cargoArea,
        situacion:$("#pcSituation").value.trim(),
        initialDate:$("#pcInitialDate").value||today(),
        finalDate:$("#pcFinalDate").value||"",
        initialObservation:$("#pcInitialObs").value.trim(),
        finalObservation:$("#pcFinalObs").value.trim(),
        technicianName:tech,
        chiefName,
        items:editedItems,
        software:editedSoftware,
        updatedAt:new Date().toISOString()
      };
  
      state.cargoOverrides[key]=payload;
  
      if(existingDoc){
        Object.assign(existingDoc,payload,{
          itemFingerprint:fingerprint,
          updatedAt:new Date().toISOString()
        });
      }
  
      await persist("EDITAR_DATOS_CARGO");
      closeModal();
      toast(`Cargo ${docCode} actualizado. El código se mantiene sin cambios.`,"success");
    };
  }
  const TIC_HEAD_NAME="FABIÁN PUENTE, FRANK JAIME";
  const CARGO_BASE_SEQUENCE=11;

  function cargoDate(value){
    const d=value?new Date(value+"T12:00:00"):new Date();
    if(Number.isNaN(d.getTime())) return value||"";
    return d.toLocaleDateString("es-PE",{day:"2-digit",month:"2-digit",year:"numeric"});
  }

  function cargoPersonKey(p){
    return p.key || (norm(p.name)+"|"+String(p.dni||"").trim());
  }

  function nextCargoCode(){
    const year=new Date().getFullYear();
    const used=new Set((state?.cargoDocuments||[])
      .map(d=>String(d.code||"").match(new RegExp("^FTEC-"+year+"-(\\d+)$","i")))
      .filter(Boolean)
      .map(m=>Number(m[1]))
      .filter(Number.isFinite));
    let seq=CARGO_BASE_SEQUENCE+1;
    while(used.has(seq)) seq++;
    return `FTEC-${year}-${String(seq).padStart(3,"0")}`;
  }

  function cargoSnapshot(items){
    return items
      .slice()
      .sort((a,b)=>(a.codigo||a.id||"").localeCompare(b.codigo||b.id||"","es"))
      .map(r=>({
        id:r.id,
        cantidad:r.cantidad||"1",
        equipo:r.equipo||"",
        bien:r.bien||"",
        marca:r.marca||"",
        modelo:r.modelo||"",
        caracteristicas:r.caracteristicas||r.detalle||r.descripcion||"",
        codigo:r.codigo||"",
        serie:r.serie||"",
        estado:r.estado||"",
        observaciones:r.observaciones||"",
        situacion:r.situacion||""
      }));
  }

  function cargoFingerprint(items){
    return (items||[]).map(r=>[
      r.id||"",r.codigo||"",r.serie||"",r.equipo||"",r.estado||""
    ].join("|")).join("||");
  }

  function cargoItemType(r){
    if(r.bien) return r.bien;
    const e=norm(r.equipo);
    return /(PC|CASE|FUENTE DE PODER|PLACA MADRE|PROCESADOR|DISCO DURO|SSD|RAM)/.test(e)?"C.INTERNO":"PERIFÉRICO";
  }

  function cargoSoftwareHtml(){
    const software=[
      ["Windows","Windows",""],["Ms Office Professional Plus","Ms Office Professional Plus",""],
      ["Winrar","Winrar",""],["VLC","VLC",""],["Aimp3","Aimp3",""],
      ["Edge","",""],["Chrome","",""],["Firefox","",""],["Adobe Acrobat","2021",""],["Anydesk","",""]
    ];
    const rows=[];
    for(let row=0;row<5;row++){
      let cells="";
      for(let block=0;block<3;block++){
        const idx=row+(block*5);
        const s=software[idx]||["","",""];
        cells+=`<td class="num">${idx+1}</td><td>${esc(s[0])}</td><td>${esc(s[1])}</td><td>${esc(s[2])}</td>`;
      }
      rows.push("<tr>"+cells+"</tr>");
    }
    return `<div class="section-label">SOFTWARE</div>
      <table class="software-table"><thead><tr>
        <th>Item</th><th>Nombre</th><th>Versión</th><th>Detalles</th>
        <th>Item</th><th>Nombre</th><th>Versión</th><th>Detalles</th>
        <th>Item</th><th>Nombre</th><th>Versión</th><th>Detalles</th>
      </tr></thead><tbody>${rows.join("")}</tbody></table>`;
  }

  function revisionBox(title,dateValue,technician,boss,observation=""){
    return `<div class="revision-box">
      <div class="revision-title">${esc(title)}</div>
      <div class="revision-date"><b>FECHA</b><span>${esc(dateValue?cargoDate(dateValue):"")}</span></div>
      <div class="revision-observation"><b>OBSERVACIONES</b><span>${esc(observation||"")}</span></div>
      <div class="revision-signatures">
        <div class="signature-cell"><div class="signature-space"></div><b>FIRMA DEL USUARIO</b></div>
        <div class="signature-cell"><div class="signature-space"></div><b>FIRMA DEL TÉCNICO</b></div>
      </div>
      <div class="revision-names">
        <div><span>${esc(technician)}</span><b>FIRMA DEL RESPONSABLE DE INF. TEC.</b></div>
        <div><span>${esc(boss)}</span><b>FIRMA DEL JEFE DE TIC</b></div>
      </div>
    </div>`;
  }

  async function printCargo(p) {
    if(!ensureEditable()) return;

    state.cargoDocuments=state.cargoDocuments||[];
    const key=cargoPersonKey(p);
    const cargoOverride=cargoOverrideFor(p);
    const cargoPersonName=(cargoOverride.personName ?? p.name ?? "").trim();
    const cargoDni=(cargoOverride.dni ?? p.dni ?? "").trim();
    const cargoSede=(cargoOverride.sede ?? [...p.sites].join(", ")).trim();
    const cargoArea=(cargoOverride.area ?? [...p.areas].join(", ")).trim();
    const editableDoc=currentCargoDocForPerson(p);
    const cargoTechnician=(cargoOverride.technicianName ?? editableDoc?.technicianName ?? "CARDENAS CURISINCHE, ROBERTO ALEJANDRO").trim();
    const cargoChief=(cargoOverride.chiefName ?? editableDoc?.chiefName ?? TIC_HEAD_NAME).trim();
    const cargoSituation=(cargoOverride.situacion ?? editableDoc?.situacion ?? "INTERNO").trim();
    const cargoInitialDate=(cargoOverride.initialDate ?? editableDoc?.initialDate ?? today()).trim();
    const cargoFinalDate=(cargoOverride.finalDate ?? editableDoc?.finalDate ?? "").trim();
    const cargoInitialObservation=(cargoOverride.initialObservation ?? editableDoc?.initialObservation ?? "").trim();
    const cargoFinalObservation=(cargoOverride.finalObservation ?? editableDoc?.finalObservation ?? "").trim();

    const baseSnap=cargoSnapshot(p.items);
    const fingerprint=cargoFingerprint(baseSnap);
    const overrideItems=Array.isArray(cargoOverride.items)&&cargoOverride.items.length?cargoOverride.items:null;
    const overrideMap=new Map((overrideItems||[]).map(x=>[x.id||x.codigo,x]));
    const snap=baseSnap.map(r=>({...r,...(overrideMap.get(r.id||r.codigo)||{})}));
    const currentDate=today();

    let doc=[...state.cargoDocuments].reverse().find(d=>
      d.personKey===key &&
      (d.itemFingerprint===fingerprint || cargoFingerprint(d.items||[])===fingerprint)
    );

    if(!doc){
      doc={
        id:crypto.randomUUID(),
        code:nextCargoCode(),
        personKey:key,
        personName:cargoPersonName,
        dni:cargoDni,
        sede:cargoSede,
        area:cargoArea,
        situacion:cargoSituation,
        initialDate:cargoInitialDate,
        finalDate:cargoFinalDate,
        initialObservation:cargoInitialObservation,
        finalObservation:cargoFinalObservation,
        technicianName:cargoTechnician,
        chiefName:cargoChief,
        software:Array.isArray(cargoOverride.software)&&cargoOverride.software.length?cargoOverride.software:cargoSoftwareDefaults(),
        itemFingerprint:fingerprint,
        items:snap,
        createdBy:currentProfile?.id||"",
        createdByName:currentProfile?.nombre||currentProfile?.email||"",
        createdAt:new Date().toISOString()
      };
      state.cargoDocuments.push(doc);
    } else {
      doc.personName=cargoPersonName;
      doc.dni=cargoDni;
      doc.sede=cargoSede;
      doc.area=cargoArea;
      doc.situacion=cargoSituation;
      doc.initialDate=cargoInitialDate;
      doc.finalDate=cargoFinalDate;
      doc.initialObservation=cargoInitialObservation;
      doc.finalObservation=cargoFinalObservation;
      doc.technicianName=cargoTechnician;
      doc.chiefName=cargoChief;
      doc.software=Array.isArray(cargoOverride.software)&&cargoOverride.software.length?cargoOverride.software:(doc.software||cargoSoftwareDefaults());
      doc.items=snap;
      doc.itemFingerprint=fingerprint;
      doc.updatedAt=new Date().toISOString();
    }

    doc.printedAt=doc.printedAt||new Date().toISOString();
    doc.lastPrintedAt=new Date().toISOString();
    doc.printCount=(Number(doc.printCount)||0)+1;
    await persist("IMPRIMIR_CARGO");

    // La plantilla oficial tiene 25 filas de hardware. Si existen más componentes,
    // se mantienen todos y se reduce proporcionalmente la altura para conservar una sola hoja.
    const items=doc.items||[];
    const slotCount=Math.max(25,items.length);
    const hardwareRows=[];
    for(let i=0;i<slotCount;i++){
      const r=items[i];
      if(r){
        const brandModel=[r.marca,r.modelo].filter(Boolean).join(" / ");
        const codeSerie=[r.codigo,r.serie].filter(Boolean).join(" / ");
        hardwareRows.push(`<tr>
          <td class="num">${i+1}</td>
          <td class="num">${esc(r.cantidad||"1")}</td>
          <td>${esc(r.equipo)}</td>
          <td>${esc(cargoItemType(r))}</td>
          <td>${esc(brandModel)}</td>
          <td>${esc(r.caracteristicas)}</td>
          <td>${esc(codeSerie||"SIN CÓDIGO")}</td>
          <td>${esc(r.estado)}</td>
          <td class="nowrap">${esc(cargoDate(doc.initialDate))}</td>
          <td></td>
          <td></td>
          <td>${esc(r.observaciones)}</td>
        </tr>`);
      }else{
        hardwareRows.push(`<tr>
          <td class="num">${i+1}</td><td></td><td></td><td></td><td></td><td></td>
          <td></td><td></td><td></td><td></td><td></td><td></td>
        </tr>`);
      }
    }

    const technician=cargoTechnician;
    const chief=cargoChief;
    const logo=document.querySelector(".brand img")?.src||document.querySelector("[data-brand-logo]")?.src||"";

    const softwareItems=Array.isArray(doc.software)&&doc.software.length?doc.software:cargoSoftwareDefaults();
    const softwareRows=[];
    for(let row=0;row<5;row++){
      const cells=[];
      for(let block=0;block<3;block++){
        const idx=row+(block*5);
        const s=softwareItems[idx]||{name:"",version:"",details:""};
        cells.push(String(idx+1),s.name||"",s.version||"",s.details||"");
      }
      softwareRows.push(`<tr>${cells.map((v,i)=>`<td class="${[0,4,8].includes(i)?"num":""}">${esc(v)}</td>`).join("")}</tr>`);
    }
    const software=softwareRows.join("");

    const reviewInitial=`
      <div class="review-box review-left">
        <div class="review-title">REVISIÓN INICIAL</div>
        <div class="review-date left-date"><b>FECHA</b><span>${esc(cargoDate(doc.initialDate))}</span></div>
        <div class="review-note-free">${esc(doc.initialObservation||"")}</div>
        <div class="review-label">OBSERVACIONES</div>
        <div class="review-write"></div>
        <div class="review-sign-labels"><b>FIRMA DEL USUARIO</b><b>FIRMA DEL TÉCNICO</b></div>
        <div class="review-sign-space"><span></span><span></span></div>
        <div class="review-names"><span>${esc(technician)}</span><span>${esc(chief)}</span></div>
        <div class="review-footer"><b>FIRMA DEL RESPONSABLE DE INF. TEC.</b><b>FIRMA DEL JEFE DE TIC</b></div>
      </div>`;

    const reviewFinal=`
      <div class="review-box review-right">
        <div class="review-title">REVISIÓN FINAL</div>
        <div class="review-date right-date"><b>FECHA</b><span>${esc(doc.finalDate?cargoDate(doc.finalDate):"")}</span></div>
        <div class="review-note-free">${esc(doc.finalObservation||"")}</div>
        <div class="review-label">OBSERVACIONES</div>
        <div class="review-write"></div>
        <div class="review-sign-labels"><b>FIRMA DEL USUARIO</b><b>FIRMA DEL TÉCNICO</b></div>
        <div class="review-sign-space"><span></span><span></span></div>
        <div class="review-names"><span>${esc(technician)}</span><span>${esc(chief)}</span></div>
        <div class="review-footer"><b>FIRMA DEL RESPONSABLE DE INF. TEC.</b><b>FIRMA DEL JEFE DE TIC</b></div>
      </div>`;

    const html=`
      <main class="excel-page">
        <div class="template-title">FICHA TÉCNICA DE EQUIPO DE CÓMPUTO</div>

        <div class="meta-wrap">
          <table class="meta-table">
            <colgroup>
              <col class="cB"><col class="cC"><col class="cD"><col class="cE"><col class="cF">
              <col class="cG"><col class="cH"><col class="cI"><col class="cJ"><col class="cK">
              <col class="cL"><col class="cM"><col class="cN"><col class="cO"><col class="cP">
            </colgroup>
            <tbody>
              <tr>
                <th colspan="2">Código:</th><td colspan="2">${esc(doc.code)}</td>
                <th>Situación</th><td colspan="3">${esc(doc.situacion||"INTERNO")}</td>
                <th>Técnico:</th><td colspan="5">${esc(technician)}</td>
                <td class="logo-cell" rowspan="2"><img src="${logo}" alt="Praxis"></td>
              </tr>
              <tr>
                <th colspan="2">Sede:</th><td colspan="2">${esc(doc.sede||"")}</td>
                <th>Área:</th><td colspan="3">${esc(doc.area||"")}</td>
                <th>Usuario:</th><td colspan="5">${esc(doc.personName)}${doc.dni?" - DNI: "+esc(doc.dni):""}</td>
              </tr>
            </tbody>
          </table>
        </div>

        <div class="section-bar">HARDWARE</div>
        <table class="hardware-table">
          <colgroup>
            <col style="width:2.542%"><col style="width:2.126%"><col style="width:10.374%"><col style="width:5.498%">
            <col style="width:10.685%"><col style="width:19.294%"><col style="width:13.693%"><col style="width:7.208%">
            <col style="width:5.705%"><col style="width:5.084%"><col style="width:4.461%"><col style="width:13.330%">
          </colgroup>
          <thead>
            <tr>
              <th>Item</th><th>Cant.</th><th>Equipo</th><th>Tipo</th><th>Marca / Modelo</th><th>Características</th>
              <th>Código TIC / Serie</th><th>Estado</th><th>Fecha</th><th>Firma (R.C.)</th><th>Firma TIC</th><th>Observación</th>
            </tr>
          </thead>
          <tbody>${hardwareRows.join("")}</tbody>
        </table>

        <div class="section-bar software-bar">SOFTWARE</div>
        <table class="software-table">
          <colgroup>
            <col style="width:2.542%"><col style="width:12.500%"><col style="width:5.498%"><col style="width:10.685%">
            <col style="width:2.437%"><col style="width:11.308%"><col style="width:5.550%"><col style="width:13.693%">
            <col style="width:2.437%"><col style="width:9.855%"><col style="width:5.084%"><col style="width:17.912%">
          </colgroup>
          <thead>
            <tr>
              <th>Ítem</th><th>Nombre</th><th>Versión</th><th>Detalles</th>
              <th>Ítem</th><th>Nombre</th><th>Versión</th><th>Detalles</th>
              <th>Ítem</th><th>Nombre</th><th>Versión</th><th>Detalles</th>
            </tr>
          </thead>
          <tbody>${software}</tbody>
        </table>

        <div class="reviews-layout">
          ${reviewInitial}
          <div class="review-gap"></div>
          ${reviewFinal}
        </div>
      </main>`;

    const w=window.open("","_blank");
    if(!w){
      toast("El navegador bloqueó la ventana de impresión. Habilita ventanas emergentes para esta página.","error");
      return;
    }

    const dynamicRowMm=slotCount>25 ? Math.max(2.15,86.3/slotCount) : 3.45;
    const dynamicFontPt=slotCount>25 ? Math.max(3.7,5.0*(25/slotCount)) : 5.0;

    w.document.write(`<!doctype html><html><head><meta charset="utf-8">
      <title>${esc(doc.code)} - ${esc(doc.personName)}</title>
      <style>
        @page{size:A4 landscape;margin:0}
        *{box-sizing:border-box}
        html,body{margin:0;padding:0;width:297mm;height:210mm;background:#fff;color:#000}
        body{font-family:Calibri,Arial,Helvetica,sans-serif;overflow:hidden}
        .excel-page{
          width:287mm;
          height:195mm;
          margin:10mm 5mm 5mm 5mm;
          overflow:hidden;
          page-break-after:avoid;
          break-after:avoid;
        }

        table{border-collapse:collapse;border-spacing:0;width:100%;table-layout:fixed}
        th,td{border:.18mm solid #000;padding:.25mm .45mm;vertical-align:middle;line-height:1.03}
        th{font-weight:700}

        .template-title{
          height:5.65mm;
          border:.18mm solid #000;
          background:#d9d9d9;
          display:flex;align-items:center;justify-content:center;
          font-size:8.6pt;font-weight:700;
          margin:0 0 3.45mm 0;
        }

        .meta-wrap{position:relative;height:7.05mm;margin-bottom:3.65mm}
        .meta-table{height:7.05mm;font-size:5.2pt}
        .meta-table tr{height:3.525mm}
        .meta-table th{background:#d9d9d9;text-align:right;padding-right:.7mm;white-space:nowrap}
        .meta-table td{background:#fff;text-align:left;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:clip}
        .meta-table .logo-cell{border:0;background:#fff;text-align:center;padding:0;vertical-align:middle}
        .meta-table .logo-cell img{max-width:31mm;max-height:11.5mm;object-fit:contain}

        .cB{width:2.542%}.cC{width:2.126%}.cD{width:10.374%}.cE{width:5.498%}.cF{width:10.685%}
        .cG{width:2.437%}.cH{width:11.308%}.cI{width:5.550%}.cJ{width:13.693%}.cK{width:2.437%}
        .cL{width:4.772%}.cM{width:5.705%}.cN{width:5.084%}.cO{width:4.461%}.cP{width:13.330%}

        .section-bar{
          height:4.13mm;
          border:.18mm solid #000;
          background:#d9d9d9;
          display:flex;align-items:center;
          font-size:5.7pt;font-weight:700;
          padding-left:.35mm;
        }

        .hardware-table thead tr{height:3.45mm}
        .hardware-table th{
          background:#d9d9d9;
          text-align:center;
          font-size:4.75pt;
          padding:.15mm .25mm;
        }
        .hardware-table tbody tr{height:${dynamicRowMm}mm}
        .hardware-table td{
          font-size:${dynamicFontPt}pt;
          padding:.12mm .35mm;
          overflow:hidden;
          white-space:normal;
        }
        .hardware-table td.num{text-align:center;font-weight:600}
        .hardware-table td:nth-child(8),
        .hardware-table td:nth-child(9){text-align:center}
        .hardware-table td:nth-child(10),
        .hardware-table td:nth-child(11){padding:0}
        .nowrap{white-space:nowrap}

        .software-bar{height:3.35mm;margin-top:1.55mm}
        .software-table thead tr{height:3.35mm}
        .software-table tbody tr{height:3.35mm}
        .software-table th{
          background:#d9d9d9;
          font-size:4.7pt;
          text-align:center;
          padding:.1mm .25mm;
        }
        .software-table td{
          font-size:4.65pt;
          padding:.1mm .35mm;
          overflow:hidden;
          white-space:nowrap;
        }
        .software-table td.num{text-align:center;font-weight:600}

        .reviews-layout{
          display:grid;
          grid-template-columns:44.969% 19.243% 35.788%;
          height:45.15mm;
          margin-top:4.0mm;
          align-items:start;
        }
        .review-gap{height:100%}
        .review-box{border:.18mm solid #000;font-size:4.85pt}
        .review-title{
          height:3.45mm;background:#d9d9d9;border-bottom:.18mm solid #000;
          display:flex;align-items:center;justify-content:center;font-weight:700;
        }
        .review-date{height:3.45mm;display:grid;border-bottom:.18mm solid #000;text-align:center}
        .left-date{grid-template-columns:45.68% 54.32%}
        .right-date{grid-template-columns:50.29% 49.71%}
        .review-date b{background:#d9d9d9;border-right:.18mm solid #000;display:flex;align-items:center;justify-content:center}
        .review-date span{display:flex;align-items:center;justify-content:center}
        .review-note-free{
          height:7.34mm;
          display:flex;align-items:center;
          padding:.3mm;
          border-bottom:.18mm solid #000;
          font-size:4.8pt;
        }
        .review-label{
          height:3.45mm;background:#d9d9d9;border-bottom:.18mm solid #000;
          display:flex;align-items:center;justify-content:center;font-weight:700;
        }
        .review-write{height:9.18mm;border-bottom:.18mm solid #000}
        .review-sign-labels{
          height:3.45mm;background:#d9d9d9;border-bottom:.18mm solid #000;
          display:grid;grid-template-columns:1fr 1fr;text-align:center
        }
        .review-sign-labels b{display:flex;align-items:center;justify-content:center}
        .review-sign-labels b:first-child{border-right:.18mm solid #000}
        .review-sign-space{
          height:9.18mm;
          display:grid;grid-template-columns:1fr 1fr;
          border-bottom:.18mm solid #000;
        }
        .review-sign-space span:first-child{border-right:.18mm solid #000}
        .review-names{
          height:3.58mm;background:#d9d9d9;
          display:grid;grid-template-columns:1fr 1fr;text-align:center;
          border-bottom:.18mm solid #000;
        }
        .review-names span{display:flex;align-items:center;justify-content:center;padding:0 .4mm;font-size:4.25pt;white-space:nowrap;overflow:hidden}
        .review-names span:first-child{border-right:.18mm solid #000}
        .review-footer{
          height:3.45mm;background:#d9d9d9;
          display:grid;grid-template-columns:1fr 1fr;text-align:center;
        }
        .review-footer b{display:flex;align-items:center;justify-content:center;font-size:4.2pt}
        .review-footer b:first-child{border-right:.18mm solid #000}

        @media print{
          html,body{width:297mm;height:210mm}
          body{-webkit-print-color-adjust:exact;print-color-adjust:exact}
          .excel-page{page-break-inside:avoid;break-inside:avoid}
        }
      </style>
    </head><body>${html}
      <script>window.onload=()=>setTimeout(()=>window.print(),250)<\/script>
    </body></html>`);
    w.document.close();
  }

  function csvDownload(records, filename, columns) {
    const lines=[columns.map(c=>`"${c[0]}"`).join(";")];
    records.forEach(r=>lines.push(columns.map(([_,key])=>`"${String(r[key]??"").replace(/"/g,'""')}"`).join(";")));
    const blob=new Blob(["\ufeff"+lines.join("\n")],{type:"text/csv;charset=utf-8"});triggerDownload(blob,filename);
  }

  function exportInventory() {
    csvDownload(inventoryFiltered(),"Inventario_TIC_Praxis.csv",[["Código TIC","codigo"],["Código Padre","codigoPadre"],["Equipo","equipo"],["Marca","marca"],["Modelo","modelo"],["Serie","serie"],["Sede","sede"],["Área","area"],["Responsable","responsable"],["DNI","dni"],["Estado","estado"],["Condición","condicion"],["Situación","situacion"],["Observaciones","observaciones"],["Origen","source"]]);
  }

  function exportMovements() {
    csvDownload([...(state.webMovements||[]),...(state.transactions||[])],"Movimientos_TIC_Praxis.csv",[["Fecha","fecha"],["Origen","source"],["Código TIC","codigo"],["Equipo","equipo"],["Desde","from"],["Hacia","to"],["Responsable","responsable"],["Observaciones","observaciones"]]);
  }

  function triggerDownload(blob,filename) {
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url;a.download=filename;a.style.display="none";
    document.body.appendChild(a);a.click();a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),1500);
  }

  function downloadBackup() {
    if(!state){toast("No hay información para generar la copia.","error");return}
    const blob=new Blob([JSON.stringify(state,null,2)],{type:"application/json;charset=utf-8"});
    triggerDownload(blob,`Backup_Inventario_Praxis_${today()}.json`);
    toast("Copia JSON generada.","success");
  }

  async function importExcel(file) {
    if(!ensureEditable()) return;
    try{
      if(state?.inventory?.length){
        const ok=confirm("Ya existe un inventario cargado en este navegador. La nueva importación reemplazará el estado actual. Se guardará una copia de recuperación automática. ¿Continuar?");
        if(!ok) return;
        await PraxisDB.set("praxis-previous-state-v2",state);
      }
      toast("Procesando Excel…");
      const next=await PraxisExcel.fromFile(file);
      if (!next.inventory.length) throw new Error("No se encontraron registros de inventario.");
      next.fileName=file.name;
      next.cargoDocuments=state?.cargoDocuments||[];
      next.cargoOverrides=state?.cargoOverrides||{};
      next.collaboratorDirectory=state?.collaboratorDirectory||[];
      state=next;await persist("IMPORTAR_EXCEL");page=1;renderAll();setView("dashboard");
      toast(`Inventario importado: ${state.inventory.length} registros maestros.`,"success");
    }catch(err){console.error(err);toast(err.message||"No se pudo importar el Excel.","error")}
  }

  async function restoreBackup(file) {
    if(!ensureEditable()) return;
    try{
      const text=await file.text();
      const restored=JSON.parse(text);
      if(!restored || !Array.isArray(restored.inventory)) throw new Error("El archivo JSON no corresponde a una copia válida del Inventario Praxis.");
      if(state?.inventory?.length && !confirm("Esta restauración reemplazará el inventario actual del navegador. ¿Continuar?")) return;
      await PraxisDB.set("praxis-previous-state-v2",state);
      state=restored;
      state.webMovements=state.webMovements||[];
      state.transactions=state.transactions||[];
      state.cargoDocuments=state.cargoDocuments||[];
      await persist("RESTAURAR_BACKUP");page=1;renderAll();setView("dashboard");
      toast("Copia de seguridad restaurada correctamente.","success");
    }catch(err){console.error(err);toast(err.message||"No se pudo restaurar la copia.","error")}
  }

  async function clearLocalData() {
    if(!ensureEditable()) return;
    if(!state?.inventory?.length) return;
    const first=confirm("Esto borrará del navegador el inventario importado y los movimientos locales. El archivo Excel original no será modificado. ¿Continuar?");
    if(!first) return;
    const second=confirm("Confirmación final: ¿borrar los datos locales de Inventario Praxis?");
    if(!second) return;
    await PraxisDB.set("praxis-previous-state-v2",state);
    await PraxisDB.remove(STATE_KEY);
    state=null;page=1;renderAll();setView("dashboard");
    toast("Datos locales borrados. Se conservó una copia de recuperación interna.","success");
  }

  function openCreateUser() {
    if(!canManageUsers()) {
      toast("Solo la cuenta Master puede crear usuarios.","error");
      return;
    }

    showModal("Crear usuario","Prepara un nuevo acceso al Inventario Praxis con DNI y rol.",`
      <div class="alert info">
        <b>Cuenta Master:</b> solo r.cardenas puede crear y administrar usuarios.
        El usuario final ingresará con su DNI.
      </div>
      <div id="createUserAlert"></div>
      <div class="form-grid">
        <div class="form-field"><label>DNI *</label><input id="cuDni" maxlength="8" inputmode="numeric" placeholder="8 dígitos"></div>
        <div class="form-field"><label>Rol *</label>
          <select id="cuRole">
            <option value="EDITOR">Editor</option>
            <option value="CONSULTA">Consulta</option>
          </select>
        </div>
        <div class="form-field full"><label>Nombres y apellidos *</label><input id="cuName"></div>
        <div class="form-field"><label>Cargo *</label><input id="cuCargo" placeholder="Ej. Secretaria, Subdirector, TIC"></div>
        <div class="form-field"><label>Sede *</label>
          <select id="cuSite">
            <option value="">Seleccionar</option>
            <option>Centro</option>
            <option>Estrellas</option>
            <option>Esperanza</option>
            <option>Estrellitas</option>
            <option>Grass Rey León</option>
            <option>Casa del Saber</option>
            <option>Bazar</option>
            <option>Otra</option>
          </select>
        </div>
        <div class="form-field full"><label>Área *</label><input id="cuArea" placeholder="Ej. Secretaría, Subdirección, TIC"></div>
      </div>
      <div class="modal-actions">
        <button id="prepareUserBtn" class="btn btn-primary">Preparar usuario</button>
      </div>
      <div id="preparedUserResult"></div>
    `);

    $("#prepareUserBtn").onclick=async()=>{
      const dni=$("#cuDni").value.trim();
      const nombre=$("#cuName").value.trim();
      const cargo=$("#cuCargo").value.trim();
      const sede=$("#cuSite").value.trim();
      const area=$("#cuArea").value.trim();
      const role=$("#cuRole").value;
      const alert=$("#createUserAlert");

      alert.className="";
      alert.textContent="";

      if(!/^\d{8}$/.test(dni)){
        alert.className="alert error";
        alert.textContent="El DNI debe tener 8 dígitos.";
        return;
      }
      if(!nombre || !cargo || !sede || !area){
        alert.className="alert error";
        alert.textContent="Completa todos los datos obligatorios.";
        return;
      }

      const btn=$("#prepareUserBtn");
      const old=btn.textContent;
      btn.disabled=true;
      btn.textContent="Preparando…";

      try{
        const internalEmail=await PraxisCloud.prepareUser({dni,nombre,cargo,sede,area,role});
        const result=$("#preparedUserResult");
        result.innerHTML=`
          <div class="alert success" style="margin-top:14px">
            <b>Usuario preparado correctamente.</b><br>
            Usuario para ingresar: <b>${esc(dni)}</b><br>
            Rol: <b>${esc(roleLabel(role))}</b><br>
            Cuenta interna de Supabase: <b>${esc(internalEmail)}</b>
          </div>
          <div class="alert info">
            Para terminar el acceso, crea este usuario una sola vez en <b>Supabase → Authentication → Users → Add user</b>
            usando la cuenta interna mostrada arriba y define su contraseña. Al crearlo, quedará autorizado automáticamente con este rol.
          </div>
          <div class="modal-actions">
            <button id="copyPreparedEmail" class="btn btn-soft">Copiar cuenta interna</button>
            <button id="openSupabaseUsers" class="btn btn-primary">Abrir Supabase Users</button>
          </div>
        `;

        $("#copyPreparedEmail").onclick=async()=>{
          try{
            await navigator.clipboard.writeText(internalEmail);
            toast("Cuenta interna copiada.","success");
          }catch(_){
            toast(internalEmail,"info");
          }
        };
        $("#openSupabaseUsers").onclick=()=>{
          window.open("https://supabase.com/dashboard/project/pgdljamqgqvfnyeuxkew/auth/users","_blank","noopener");
        };

        toast("Usuario preparado por la cuenta Master.","success");
      }catch(err){
        console.error(err);
        alert.className="alert error";
        alert.textContent=(err&&err.message)||"No se pudo preparar el usuario.";
      }finally{
        btn.disabled=false;
        btn.textContent=old;
      }
    };
  }

  async function renderUsers() {
    if(!canManageUsers() || !$("#usersTable")) return;
    try {
      const allUsers=await PraxisCloud.listUsers();
      const q=norm($("#authorizedSearch")?.value||"");
      const status=$("#authorizedStatus")?.value||"";
      const role=$("#authorizedRole")?.value||"";

      const activeCount=allUsers.filter(u=>u.activo).length;
      const pendingCount=allUsers.length-activeCount;
      const admins=allUsers.filter(u=>u.role==="ADMIN_TIC"&&u.activo).length;
      const editors=allUsers.filter(u=>u.role==="EDITOR"&&u.activo).length;
      const consult=allUsers.filter(u=>u.role==="CONSULTA"&&u.activo).length;
      const printedDocs=(state?.cargoDocuments||[]).filter(d=>d.printedAt);
      const signedDocs=printedDocs.filter(d=>d.receivedSigned);

      if($("#authorizedSummary")) $("#authorizedSummary").innerHTML=[
        [allUsers.length,"Cuentas registradas"],
        [activeCount,"Personal autorizado"],
        [pendingCount,"Pendientes / bloqueados"],
        [admins,"Cuenta Master"],
        [editors,"Editores"],
        [consult,"Solo lectura"],
        [printedDocs.length,"Cargos impresos"],
        [signedDocs.length,"Firmados / recibidos"]
      ].map(([n,l])=>`<div class="quality-box"><b>${n}</b><span>${esc(l)}</span></div>`).join("");

      const users=allUsers.filter(u=>{
        const hay=norm([u.nombre,u.email,u.dni,u.cargo,u.sede,u.area,u.role].join(" "));
        return (!q||hay.includes(q))
          && (!status||(status==="ACTIVO"?u.activo:!u.activo))
          && (!role||u.role===role);
      });

      const me=currentProfile?.id;
      const rows=users.map(u=>{
        const isMe=u.id===me;
        const isMaster=String(u.email||"").toLowerCase()===OWNER_ADMIN_EMAIL;
        const statusBadge=u.activo
          ? '<span class="badge ok">AUTORIZADO</span>'
          : '<span class="badge warn">PENDIENTE / BLOQUEADO</span>';
        const actionText=u.activo?"Desautorizar":"Autorizar";
        const actionClass=u.activo?"btn-danger":"btn-success";

        const userDni=String(u.dni||"").replace(/\D/g,"");
        const userCargoDocs=printedDocs.filter(d=>userDni && String(d.dni||"").replace(/\D/g,"")===userDni);
        const latestCargo=[...userCargoDocs].reverse()[0]||null;
        const cargoControl=latestCargo
          ? `<div class="cargo-user-cell"><b>${esc(latestCargo.code)}</b><div class="muted">${userCargoDocs.length} cargo(s) impreso(s)</div>
              <span class="badge ${latestCargo.receivedSigned?"ok":"warn"}">${latestCargo.receivedSigned?"FIRMADO / RECIBIDO":"PENDIENTE FIRMA"}</span>
              <button class="btn ${latestCargo.receivedSigned?"btn-warning":"btn-success"} btn-small" data-received-cargo="${esc(latestCargo.id)}" data-next-received="${latestCargo.receivedSigned?"false":"true"}">${latestCargo.receivedSigned?"Quitar recibido":"Marcar recibido"}</button>
            </div>`
          : '<span class="muted">Sin cargo impreso</span>';

        const roleControl=isMaster
          ? '<select class="role-select" disabled><option selected>MASTER TIC</option></select>'
          : `<select class="role-select" data-role-user="${esc(u.id)}">
              <option value="EDITOR" ${u.role==="EDITOR"?"selected":""}>Editor</option>
              <option value="CONSULTA" ${u.role==="CONSULTA"?"selected":""}>Solo lectura</option>
            </select>`;

        return `<tr>
          <td>
            <div class="user-name-cell"><span class="avatar mini">TIC</span><div><b>${esc(u.nombre||"Sin nombre")}</b>${isMaster?'<span class="you-tag">MASTER</span>':""}<div class="muted">${isMaster?esc(u.email||""):"Usuario DNI: "+esc(u.dni||"—")}</div><div class="muted">DNI: ${esc(u.dni||"—")}</div></div></div>
          </td>
          <td><b>${esc(u.cargo||"—")}</b><div class="muted">${esc([u.sede,u.area].filter(Boolean).join(" / ")||"—")}</div></td>
          <td>${roleControl}</td>
          <td>${statusBadge}</td>
          <td>${cargoControl}</td>
          <td>${esc(fmtDateTime(u.last_login_at))}</td>
          <td><button class="btn ${actionClass} btn-small" data-active-user="${esc(u.id)}" data-next-active="${u.activo?"false":"true"}" ${isMaster?"disabled":""}>${actionText}</button></td>
        </tr>`;
      }).join("");

      $("#usersTable").innerHTML=`<div class="table-wrap"><table class="authorized-table"><thead><tr><th>Personal</th><th>Cargo / ubicación</th><th>Rol</th><th>Acceso</th><th>Cargos / recepción</th><th>Último acceso</th><th>Acción</th></tr></thead><tbody>${rows||'<tr><td colspan="7">No se encontró personal con esos filtros.</td></tr>'}</tbody></table></div>`;

      $$("[data-role-user]").forEach(sel=>{
        sel.onchange=async()=>{
          const role=sel.value;
          if(!confirm("¿Cambiar el rol de este usuario a "+roleLabel(role)+"?")){await renderUsers();return;}
          try{
            await PraxisCloud.setUserRole(sel.dataset.roleUser,role);
            toast("Rol actualizado.","success");
            await renderUsers();
            await renderAudit();
          }catch(err){
            console.error(err);
            toast((err&&err.message)||"No se pudo actualizar el rol.","error");
            await renderUsers();
          }
        };
      });

      $("[data-received-cargo]").forEach(btn=>{
        btn.onclick=async()=>{
          const received=btn.dataset.nextReceived==="true";
          if(!confirm(received?"¿Confirmar que el usuario firmó el cargo y recibió sus equipos?":"¿Quitar la confirmación de recepción?")) return;
          try{
            await setCargoReceived(btn.dataset.receivedCargo,received);
            await renderUsers();
            await renderAudit();
          }catch(err){
            console.error(err);
            toast((err&&err.message)||"No se pudo actualizar la recepción del cargo.","error");
          }
        };
      });

      $("[data-active-user]").forEach(btn=>{
        btn.onclick=async()=>{
          const active=btn.dataset.nextActive==="true";
          const label=active?"autorizar":"desautorizar";
          if(!confirm("¿Seguro que deseas "+label+" a este usuario?")) return;
          try{
            await PraxisCloud.setUserActive(btn.dataset.activeUser,active);
            toast(active?"Personal autorizado.":"Acceso retirado.","success");
            await renderUsers();
            await renderAudit();
          }catch(err){
            console.error(err);
            toast((err&&err.message)||"No se pudo actualizar la autorización.","error");
          }
        };
      });
    } catch(err) {
      console.error(err);
      const table=$("#usersTable");
      if(table && !table.querySelector("table")){
        table.innerHTML="<div class='alert error'>"+esc((err&&err.message)||"No se pudo cargar el personal autorizado.")+"</div>";
      } else {
        toast((err&&err.message)||"Ocurrió un error al enlazar los controles de usuarios.","error");
      }
    }
  }

  function cargoAuditDocs(){
    return (state?.cargoDocuments||[])
      .filter(d=>d.printedAt)
      .slice()
      .sort((a,b)=>String(b.printedAt||b.createdAt||"").localeCompare(String(a.printedAt||a.createdAt||"")));
  }
  
  function renderCargoAudit(){
    if(!canManageUsers() || !$("#cargoAuditTable")) return;
  
    const all=cargoAuditDocs();
    const site=$("#cargoAuditSite")?.value||"";
    const status=$("#cargoAuditStatus")?.value||"";
    const q=norm($("#cargoAuditSearch")?.value||"");
  
    const sites=[...new Set(all.map(d=>String(d.sede||"").trim()).filter(Boolean))]
      .sort((a,b)=>a.localeCompare(b,"es"));
  
    const siteSelect=$("#cargoAuditSite");
    if(siteSelect){
      const current=siteSelect.value;
      siteSelect.innerHTML='<option value="">Todas las sedes</option>'+
        sites.map(s=>`<option value="${esc(s)}">${esc(s)}</option>`).join("");
      if(sites.includes(current)) siteSelect.value=current;
    }
  
    const signed=all.filter(d=>d.receivedSigned);
    const pending=all.filter(d=>!d.receivedSigned);
    const reprints=all.reduce((sum,d)=>sum+Math.max(0,(Number(d.printCount)||1)-1),0);
  
    if($("#cargoAuditSummary")){
      $("#cargoAuditSummary").innerHTML=[
        [all.length,"Cargos impresos"],
        [signed.length,"Firmados / recibidos"],
        [pending.length,"Pendientes de firma"],
        [sites.length,"Sedes con cargos"],
        [reprints,"Reimpresiones"]
      ].map(([n,l])=>`<div class="quality-box"><b>${n}</b><span>${esc(l)}</span></div>`).join("");
    }
  
    const bySite=new Map();
    all.forEach(d=>{
      const s=String(d.sede||"SIN SEDE").trim()||"SIN SEDE";
      if(!bySite.has(s)) bySite.set(s,{total:0,signed:0,pending:0});
      const x=bySite.get(s);
      x.total++;
      if(d.receivedSigned)x.signed++;else x.pending++;
    });
  
    if($("#cargoAuditBySite")){
      $("#cargoAuditBySite").innerHTML=[...bySite.entries()]
        .sort((a,b)=>b[1].total-a[1].total || a[0].localeCompare(b[0],"es"))
        .map(([s,v])=>`
          <button class="cargo-site-card ${site===s?"active":""}" type="button" data-cargo-site="${esc(s)}">
            <span>${esc(s)}</span>
            <b>${v.total}</b>
            <small>${v.signed} firmados · ${v.pending} pendientes</small>
          </button>`).join("") || '<div class="muted">Todavía no hay cargos impresos.</div>';
      $$("[data-cargo-site]").forEach(b=>b.onclick=()=>{
        if($("#cargoAuditSite")) $("#cargoAuditSite").value=b.dataset.cargoSite;
        renderCargoAudit();
      });
    }
  
    const docs=all.filter(d=>{
      const matchesSite=!site || String(d.sede||"")===site;
      const matchesStatus=!status ||
        (status==="SIGNED" && d.receivedSigned) ||
        (status==="PENDING" && !d.receivedSigned);
      const hay=norm([d.code,d.personName,d.dni,d.sede,d.area,d.createdByName,d.receivedByName].join(" "));
      return matchesSite && matchesStatus && (!q || hay.includes(q));
    });
  
    if($("#cargoAuditResultCount")){
      $("#cargoAuditResultCount").innerHTML=`Mostrando <b>${docs.length}</b> de <b>${all.length}</b> cargos impresos`;
    }
  
    const rows=docs.map(d=>`
      <tr>
        <td><span class="code">${esc(d.code||"—")}</span><div class="muted">Impresiones: ${Number(d.printCount)||1}</div></td>
        <td><b>${esc(d.personName||"Sin nombre")}</b><div class="muted">DNI: ${esc(d.dni||"—")}</div></td>
        <td><b>${esc(d.sede||"—")}</b><div class="muted">${esc(d.area||"—")}</div></td>
        <td>${esc(fmtDateTime(d.printedAt||d.createdAt))}</td>
        <td>${d.receivedSigned
          ? `<span class="badge ok">FIRMADO / RECIBIDO</span><div class="muted">${esc(fmtDateTime(d.receivedAt))}</div>`
          : '<span class="badge warn">PENDIENTE DE FIRMA</span>'}</td>
        <td>${esc(d.createdByName||"—")}</td>
        <td>
          <div class="button-row">
            <button class="btn btn-soft btn-small" type="button" data-view-cargo="${esc(d.id)}">Ver detalle</button>
            <button class="btn ${d.receivedSigned?"btn-warning":"btn-success"} btn-small" type="button" data-audit-received="${esc(d.id)}" data-next-received="${d.receivedSigned?"false":"true"}">${d.receivedSigned?"Quitar recibido":"Marcar recibido"}</button>
          </div>
        </td>
      </tr>`).join("");
  
    $("#cargoAuditTable").innerHTML=`<div class="table-wrap cargo-audit-table"><table>
      <thead><tr><th>Código FTEC</th><th>Colaborador</th><th>Sede / área</th><th>Fecha impresión</th><th>Estado de recepción</th><th>Impreso por</th><th>Acciones</th></tr></thead>
      <tbody>${rows||'<tr><td colspan="7">No hay cargos que coincidan con los filtros seleccionados.</td></tr>'}</tbody>
    </table></div>`;
  
    $$("[data-audit-received]").forEach(btn=>btn.onclick=async()=>{
      const received=btn.dataset.nextReceived==="true";
      if(!confirm(received?"¿Confirmar que el usuario firmó el cargo y recibió los equipos?":"¿Quitar la confirmación de recepción?")) return;
      await setCargoReceived(btn.dataset.auditReceived,received);
      renderCargoAudit();
      await renderUsers();
      await renderAudit();
    });
  
    $$("[data-view-cargo]").forEach(btn=>btn.onclick=()=>{
      const d=all.find(x=>x.id===btn.dataset.viewCargo);
      if(!d)return;
      showModal(`Cargo ${d.code||""}`,d.personName||"Detalle del cargo",`
        <div class="detail-grid">
          <div class="detail-field"><span>Colaborador</span><b>${esc(d.personName||"—")}</b></div>
          <div class="detail-field"><span>DNI</span><b>${esc(d.dni||"—")}</b></div>
          <div class="detail-field"><span>Sede</span><b>${esc(d.sede||"—")}</b></div>
          <div class="detail-field"><span>Área</span><b>${esc(d.area||"—")}</b></div>
          <div class="detail-field"><span>Fecha impresión</span><b>${esc(fmtDateTime(d.printedAt||d.createdAt))}</b></div>
          <div class="detail-field"><span>Estado</span><b>${d.receivedSigned?"FIRMADO / RECIBIDO":"PENDIENTE DE FIRMA"}</b></div>
        </div>
        <h4 class="section-title">Equipos incluidos (${(d.items||[]).length})</h4>
        <div class="table-wrap"><table><thead><tr><th>Código TIC</th><th>Equipo</th><th>Marca / modelo</th><th>Serie</th><th>Estado</th></tr></thead>
        <tbody>${(d.items||[]).map(r=>`<tr><td><span class="code">${esc(r.codigo||"—")}</span></td><td>${esc(r.equipo||"—")}</td><td>${esc([r.marca,r.modelo].filter(Boolean).join(" / ")||"—")}</td><td>${esc(r.serie||"—")}</td><td>${esc(r.estado||"—")}</td></tr>`).join("")||'<tr><td colspan="5">Sin equipos registrados.</td></tr>'}</tbody></table></div>
      `);
    });
  }

  async function renderAudit() {
    if(!canManageUsers() || !$("#auditTable")) return;
    try {
      const [items,users]=await Promise.all([PraxisCloud.audit(100),PraxisCloud.listUsers()]);
      const userMap=new Map(users.map(u=>[u.id,u]));
      const actionLabel=action=>({
        IMPORTAR_EXCEL:"Importación de Excel",
        ALTA_EQUIPO:"Alta de equipo",
        EDITAR_FICHA:"Edición de ficha",
        EDITAR_COLABORADOR:"Edición de datos de cargo",
        CAMBIAR_COLABORADOR_RESPONSABLE:"Cambio de colaborador / responsable",
        QUITAR_COLABORADOR:"Retiro de colaborador",
        EDITAR_DATOS_CARGO:"Edición completa de ficha de cargo",
        IMPRIMIR_CARGO:"Impresión / reimpresión de cargo",
        CONFIRMAR_RECEPCION_CARGO:"Confirmación de firma y recepción",
        MOVIMIENTO_EQUIPO:"Movimiento / transferencia",
        RESOLVER_CONFLICTO:"Resolución de conflicto",
        RESTAURAR_BACKUP:"Restauración de copia",
        IMPRIMIR_CARGO:"Ficha técnica · cargo de equipos",
        CAMBIAR_ROL:"Cambio de rol",
        AUTORIZAR_USUARIO:"Autorización de personal",
        DESAUTORIZAR_USUARIO:"Retiro de acceso",
        BOOTSTRAP_ADMIN:"Alta de cuenta Master",
        CREAR_USUARIO:"Creación de usuario",
        }[action]||action||"Evento");

      const rows=items.map(a=>{
        const u=userMap.get(a.user_id);
        const person=u?(u.nombre||u.email):a.user_id||"Sistema";
        const summary=a.summary||{};
        const detail=[
          summary.fileName,
          summary.masterRecords!=null?summary.masterRecords+" registros":null,
          summary.new_role?"Rol: "+summary.new_role:null
        ].filter(Boolean).join(" · ");
        return `<tr>
          <td>${esc(fmtDateTime(a.created_at))}</td>
          <td><b>${esc(actionLabel(a.action))}</b>${detail?'<div class="muted">'+esc(detail)+'</div>':""}</td>
          <td>${esc(person)}</td>
          <td>${esc(a.version_before==null?"—":a.version_before)}</td>
          <td>${esc(a.version_after==null?"—":a.version_after)}</td>
        </tr>`;
      }).join("");
      $("#auditTable").innerHTML=`<div class="table-wrap"><table><thead><tr><th>Fecha</th><th>Acción</th><th>Realizado por</th><th>Versión anterior</th><th>Nueva versión</th></tr></thead><tbody>${rows||'<tr><td colspan="5">Sin eventos de auditoría.</td></tr>'}</tbody></table></div>`;
    } catch(err) {
      console.error(err);
      $("#auditTable").innerHTML="<div class='alert error'>"+esc((err&&err.message)||"No se pudo cargar la auditoría.")+"</div>";
    }
  }

  async function enterCloudApp() {
    currentProfile=await PraxisCloud.loadProfile();
    if(!currentProfile) throw new Error("No se encontró el perfil institucional.");

    if(!currentProfile.activo || currentProfile.role==="CONSULTA"){
      try{
        const promoted=await PraxisCloud.bootstrapAdmin();
        if(promoted) currentProfile=await PraxisCloud.loadProfile();
      }catch(err){ console.warn("Bootstrap admin no aplicado",err); }
    }

    if(!currentProfile.activo){
      throw new Error("Tu cuenta está pendiente de autorización por el Administrador TIC.");
    }

    const remote=await PraxisCloud.loadState();
    state=remote.state;

    // La sesión ya fue validada por Supabase. Desde este punto ningún fallo
    // de caché local o tiempo real debe devolver al usuario al login.
    $("#authGate")?.classList.add("hidden");
    $("#appShell")?.classList.remove("hidden");
    updateProfileUI();

    try{
      if(state) await PraxisDB.set(STATE_KEY,state);
    }catch(err){
      console.warn("No se pudo guardar la caché local; se continuará con Supabase.",err);
    }

    try{
      if(realtimeChannel) await PraxisCloud.unsubscribe(realtimeChannel);
      realtimeChannel=PraxisCloud.subscribeState(async latest=>{
        state=latest.state;
        try{ if(state) await PraxisDB.set(STATE_KEY,state); }catch(err){ console.warn("Caché local no disponible",err); }
        try{ renderAll(); }catch(err){ console.error("Error al refrescar inventario",err); }
      });
    }catch(err){
      console.warn("Tiempo real no disponible; el sistema seguirá funcionando.",err);
      realtimeChannel=null;
    }

    try{
      if(accessChannel) await PraxisCloud.unsubscribe(accessChannel);
      accessChannel=PraxisCloud.subscribeProfile(async latestProfile=>{
        const previousRole=currentProfile?.role;
        const previousActive=currentProfile?.activo;
        currentProfile=latestProfile;

        if(!currentProfile?.activo){
          try{
            if(realtimeChannel) await PraxisCloud.unsubscribe(realtimeChannel);
            if(accessChannel) await PraxisCloud.unsubscribe(accessChannel);
          }catch(_){}
          realtimeChannel=null;
          accessChannel=null;
          await PraxisCloud.signOut();
          state=null;
          showLogin();
          return;
        }

        updateProfileUI();

        if(previousRole!==currentProfile.role || previousActive!==currentProfile.activo){
          if(!canEdit()){
            closeModal();
            setView("dashboard");
          }
          renderAll();
          toast(currentProfile.role==="CONSULTA"
            ? "Tu acceso cambió a Solo lectura."
            : "Tus permisos de acceso fueron actualizados.","success");
        }
      });
    }catch(err){
      console.warn("No se pudo sincronizar el rol del usuario en tiempo real.",err);
      accessChannel=null;
    }

    try{
      renderAll();
      setView("dashboard");
    }catch(err){
      console.error("Error al pintar el panel",err);
      const workspace=$("#workspace");
      const empty=$("#emptyState");
      if(workspace) workspace.classList.remove("hidden");
      if(empty) empty.classList.add("hidden");
      toast("Sesión iniciada. Ocurrió un problema al mostrar una sección del panel; recarga la página.","error");
    }
  }

  function showAuth(message,type) {
    $("#appShell").classList.add("hidden");
    $("#authGate").classList.remove("hidden");
    const box=$("#authMessage");
    if(message){
      box.textContent=message;
      box.className="auth-message show "+(type||"info");
    } else {
      box.textContent="";
      box.className="auth-message";
    }
  }

  function showLogin(message="",type="info") {
    const form=$("#loginForm");
    const subtitle=$("#authSubtitle");
    if(form) form.classList.remove("hidden");
    if(subtitle) subtitle.textContent="Acceso exclusivo para personal autorizado por TIC.";
    showAuth(message,type);
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

  function wireAuthControls() {
    const loginForm = document.querySelector("#loginForm");

    if (loginForm) {
      loginForm.onsubmit = async (e) => {
        e.preventDefault();

        const user = document.querySelector("#loginEmail")?.value?.trim() || "";
        const password = document.querySelector("#loginPassword")?.value || "";
        const box = document.querySelector("#authMessage");
        const button = loginForm.querySelector("button[type='submit']");

        if (!user || !password) {
          if (box) {
            box.textContent = "Ingrese usuario y contraseña";
            box.className = "auth-message show error";
          }
          return;
        }

        if (button) {
          button.disabled = true;
          button.textContent = "Iniciando...";
        }

        if (box) {
          box.textContent = "";
          box.className = "auth-message";
        }

        try {
          await PraxisCloud.signIn(user, password);

          if (button) {
            button.disabled = false;
            button.textContent = "Iniciar sesión";
          }

          if (box) {
            box.textContent = "Sesión iniciada correctamente, redireccionando...";
            box.className = "auth-message show success";
          }

          await new Promise(resolve => setTimeout(resolve, 1000));
          await enterCloudApp();

        } catch (err) {
          console.error("Login falló", err);

          if (button) {
            button.disabled = false;
            button.textContent = "Iniciar sesión";
          }

          if (box) {
            box.textContent = "La contraseña ingresada es incorrecta";
            box.className = "auth-message show error";
          }

          const passwordField = document.querySelector("#loginPassword");
          if (passwordField) {
            passwordField.focus();
            passwordField.select();
          }
        }
      };
    }
  }

  /* COLLABORATOR_ASSIGNMENT_V2 */
function collaboratorIdentity(name,dni) {
  return `NOMBRE:${norm(name)}`;
}

function upsertCollaboratorDirectory(person) {
  state.collaboratorDirectory=state.collaboratorDirectory||[];
  const name=String(person?.name||'').trim();
  const dni=String(person?.dni||'').replace(/\D/g,'').slice(0,8);
  if(!name) return null;
  const idx=state.collaboratorDirectory.findIndex(c=>norm(c.name)===norm(name));
  const base={
    id:idx>=0 ? state.collaboratorDirectory[idx].id : crypto.randomUUID(),
    name,dni,
    sede:String(person?.sede||'').trim().toUpperCase(),
    area:String(person?.area||'').trim(),
    updatedAt:new Date().toISOString()
  };
  if(idx>=0){
    state.collaboratorDirectory[idx]={...state.collaboratorDirectory[idx],...base};
    return state.collaboratorDirectory[idx];
  }
  base.createdAt=new Date().toISOString();
  state.collaboratorDirectory.push(base);
  return base;
}

function collaborators() {
  const map=new Map();
  const ensure=name=>{
    const clean=String(name||'').trim();
    const n=norm(clean);
    if(!n) return null;
    const key=`NOMBRE:${n}`;
    if(!map.has(key)) map.set(key,{key,name:clean,dni:'',items:[],sites:new Set(),areas:new Set(),manual:false,manualDni:'',dniCounts:new Map()});
    return map.get(key);
  };
  const addDni=(p,dni,manual=false)=>{
    const clean=String(dni||'').replace(/\D/g,'').slice(0,8);
    if(!clean) return;
    if(manual) p.manualDni=clean;
    p.dniCounts.set(clean,(p.dniCounts.get(clean)||0)+1);
  };
  (state?.collaboratorDirectory||[]).forEach(c=>{
    const p=ensure(c.name); if(!p) return;
    p.manual=true; addDni(p,c.dni,true);
    if(c.sede) p.sites.add(c.sede); if(c.area) p.areas.add(c.area);
  });
  activeInventory().filter(r=>String(r.responsable||'').trim()).forEach(r=>{
    const p=ensure(r.responsable); if(!p) return;
    p.items.push(r); addDni(p,r.dni,false);
    if(r.sede) p.sites.add(r.sede); if(r.area) p.areas.add(r.area);
  });
  for(const p of map.values()){
    if(p.manualDni) p.dni=p.manualDni;
    else p.dni=[...p.dniCounts.entries()].sort((x,y)=>y[1]-x[1]||x[0].localeCompare(y[0]))[0]?.[0]||'';
    delete p.manualDni; delete p.dniCounts;
  }
  return [...map.values()].sort((x,y)=>x.name.localeCompare(y.name,'es'));
}
function collaboratorRecordMatches(r,p){
  if(!r || !p || !String(r.responsable||"").trim()) return false;
  const rdni=String(r.dni||"").replace(/\D/g,"");
  const pdni=String(p.dni||"").replace(/\D/g,"");
  return Boolean(rdni && pdni && rdni===pdni) || norm(r.responsable)===norm(p.name);
}

function collaboratorDirectoryMatches(c,p){
  if(!c || !p) return false;
  const cdni=String(c.dni||"").replace(/\D/g,"");
  const pdni=String(p.dni||"").replace(/\D/g,"");
  return Boolean(cdni && pdni && cdni===pdni) || norm(c.name)===norm(p.name);
}

function openEditCollaborator(p){
  if(!ensureEditable()) return;

  const equipmentCount=p.items?.length||0;
  const sites=[...p.sites].join(", ")||"—";
  const areas=[...p.areas].join(", ")||"—";

  showModal("Editar colaborador","Cambia el nombre o DNI del responsable sin mover ni perder sus equipos.",`
    <div class="alert info">
      <b>Se mantendrán los mismos ${equipmentCount} equipo(s).</b>
      Solo cambiará el responsable. La sede, área, Código TIC, serie y demás datos de los equipos permanecen igual.
    </div>
    <div id="editCollaboratorAlert"></div>
    <div class="detail-grid">
      <div class="detail-field"><span>Sede(s) actual(es)</span><b>${esc(sites)}</b></div>
      <div class="detail-field"><span>Área(s) actual(es)</span><b>${esc(areas)}</b></div>
      <div class="detail-field"><span>Equipos que conservará</span><b>${equipmentCount}</b></div>
    </div>
    <div class="form-grid collaborator-edit-form">
      <div class="form-field full"><label>Nombres y apellidos *</label><input id="editCollaboratorName" value="${esc(p.name||"")}"></div>
      <div class="form-field full"><label>DNI</label><input id="editCollaboratorDni" maxlength="8" inputmode="numeric" value="${esc(p.dni||"")}" placeholder="8 dígitos"></div>
    </div>
    <div class="modal-actions">
      <button id="saveCollaboratorProfile" class="btn btn-primary">Guardar nuevo responsable</button>
    </div>
  `);

  $("#saveCollaboratorProfile").onclick=async()=>{
    const alert=$("#editCollaboratorAlert");
    const newName=$("#editCollaboratorName").value.trim();
    const newDni=$("#editCollaboratorDni").value.replace(/\D/g,"").slice(0,8);
    alert.className="";
    alert.textContent="";

    if(!newName){
      alert.className="alert error";
      alert.textContent="Ingresa el nombre del colaborador.";
      return;
    }
    if(newDni && !/^\d{8}$/.test(newDni)){
      alert.className="alert error";
      alert.textContent="El DNI debe tener 8 dígitos.";
      return;
    }

    const duplicate=collaborators().find(x=>{
      if(x.key===p.key) return false;
      const xDni=String(x.dni||"").replace(/\D/g,"");
      return (newDni && xDni===newDni) || norm(x.name)===norm(newName);
    });
    if(duplicate){
      alert.className="alert error";
      alert.textContent=`Ya existe otro colaborador registrado como ${duplicate.name}${duplicate.dni?` (DNI ${duplicate.dni})`:""}.`;
      return;
    }

    const oldName=p.name||"";
    const oldDni=p.dni||"";
    const affected=activeInventory().filter(r=>collaboratorRecordMatches(r,p));
    const firstSite=[...p.sites][0]||affected[0]?.sede||"";
    const firstArea=[...p.areas][0]||affected[0]?.area||"";

    state.collaboratorDirectory=state.collaboratorDirectory||[];
    state.webMovements=state.webMovements||[];

    state.collaboratorDirectory=state.collaboratorDirectory.filter(c=>!collaboratorDirectoryMatches(c,p));
    upsertCollaboratorDirectory({name:newName,dni:newDni,sede:firstSite,area:firstArea});

    affected.forEach(r=>{
      r.history=r.history||[];
      r.history.unshift({
        type:"WEB",
        source:"Cambio de colaborador",
        fecha:today(),
        sede:r.sede||"",
        area:r.area||"",
        responsable:oldName,
        estado:r.estado||"",
        observaciones:`Cambio de responsable: ${oldName}${oldDni?` (DNI ${oldDni})`:""} → ${newName}${newDni?` (DNI ${newDni})`:""}. El equipo permanece en la misma ubicación.`
      });

      state.webMovements.unshift({
        id:crypto.randomUUID(),
        fecha:today(),
        source:"Cambio de colaborador",
        codigo:r.codigo,
        equipo:r.equipo,
        from:[r.sede,r.area,oldName].filter(Boolean).join(" / "),
        to:[r.sede,r.area,newName].filter(Boolean).join(" / "),
        responsable:newName,
        motivo:"Cambio de colaborador manteniendo los mismos equipos",
        observaciones:`Responsable anterior: ${oldName}. Nuevo responsable: ${newName}.`
      });

      r.responsable=newName;
      r.dni=newDni;
      if(r.locationType!=="ALMACEN") r.locationType="ASIGNADO";
    });

    await persist("CAMBIAR_COLABORADOR_RESPONSABLE");
    closeModal();
    renderAll();
    if($("#peopleSearch")){
      $("#peopleSearch").value=newDni||newName;
      renderPeople();
    }
    toast(`${newName}: ${affected.length} equipo(s) conservados a su nombre.`,"success");
  };
}

function openRemoveCollaborator(p){
  if(!ensureEditable()) return;
  const equipmentCount=p.items?.length||0;

  showModal("Quitar colaborador","Retira al colaborador sin borrar ningún equipo del inventario.",`
    <div class="alert error">
      <b>Esta acción quitará a ${esc(p.name)} como responsable.</b>
      Los ${equipmentCount} equipo(s) seguirán registrados con su mismo Código TIC, sede y área, pero quedarán <b>sin responsable</b>.
      Los cargos ya impresos se conservan como historial.
    </div>
    <div class="detail-grid">
      <div class="detail-field"><span>Colaborador</span><b>${esc(p.name||"—")}</b></div>
      <div class="detail-field"><span>DNI</span><b>${esc(p.dni||"—")}</b></div>
      <div class="detail-field"><span>Equipos afectados</span><b>${equipmentCount}</b></div>
    </div>
    <div class="modal-actions">
      <button id="confirmRemoveCollaborator" class="btn btn-danger">Quitar colaborador</button>
    </div>
  `);

  $("#confirmRemoveCollaborator").onclick=async()=>{
    if(!confirm(`¿Seguro que deseas quitar a ${p.name} como responsable de ${equipmentCount} equipo(s)?`)) return;

    const oldName=p.name||"";
    const oldDni=p.dni||"";
    const affected=activeInventory().filter(r=>collaboratorRecordMatches(r,p));

    state.collaboratorDirectory=state.collaboratorDirectory||[];
    state.webMovements=state.webMovements||[];
    state.collaboratorDirectory=state.collaboratorDirectory.filter(c=>!collaboratorDirectoryMatches(c,p));

    affected.forEach(r=>{
      r.history=r.history||[];
      r.history.unshift({
        type:"WEB",
        source:"Retiro de colaborador",
        fecha:today(),
        sede:r.sede||"",
        area:r.area||"",
        responsable:oldName,
        estado:r.estado||"",
        observaciones:`Se retiró al responsable ${oldName}${oldDni?` (DNI ${oldDni})`:""}. El equipo permanece en su misma sede y área sin responsable.`
      });

      state.webMovements.unshift({
        id:crypto.randomUUID(),
        fecha:today(),
        source:"Retiro de colaborador",
        codigo:r.codigo,
        equipo:r.equipo,
        from:[r.sede,r.area,oldName].filter(Boolean).join(" / "),
        to:[r.sede,r.area,"SIN RESPONSABLE"].filter(Boolean).join(" / "),
        responsable:"",
        motivo:"Retiro de colaborador",
        observaciones:`Se quitó a ${oldName} como responsable. El equipo conserva su ubicación.`
      });

      r.responsable="";
      r.dni="";
      if(r.locationType==="ASIGNADO") r.locationType="SEDE";
    });

    await persist("QUITAR_COLABORADOR");
    closeModal();
    if($("#peopleSearch")) $("#peopleSearch").value="";
    renderAll();
    toast(`${oldName} fue retirado. ${affected.length} equipo(s) quedaron sin responsable y conservaron su ubicación.`,"success");
  };
}

function renderPeople(){
  const search=$("#peopleSearch");
  const raw=String(search?.value||"").trim();
  const q=norm(raw);
  const digits=raw.replace(/\D/g,"");
  const tokens=q.split(/\s+/).filter(Boolean);
  const all=collaborators();

  const scored=all.map(p=>{
    const name=norm(p.name||"");
    const dni=String(p.dni||"").replace(/\D/g,"");
    const siteArea=norm([...p.sites,...p.areas].join(" "));
    let score=0;
    let match=!raw;

    if(raw){
      if(digits){
        if(dni===digits){score=1000;match=true}
        else if(dni.startsWith(digits)){score=850;match=true}
        else if(dni.includes(digits)){score=700;match=true}
      }
      if(q){
        if(name===q){score=Math.max(score,950);match=true}
        else if(name.startsWith(q)){score=Math.max(score,800);match=true}
        else if(name.includes(q)){score=Math.max(score,650);match=true}
        else if(tokens.length && tokens.every(t=>name.includes(t))){score=Math.max(score,600);match=true}
        else if(siteArea.includes(q)){score=Math.max(score,200);match=true}
      }
    }
    return {p,score,match};
  }).filter(x=>x.match)
    .sort((a,b)=>b.score-a.score || a.p.name.localeCompare(b.p.name,"es"));

  const ps=scored.map(x=>x.p);
  const status=$("#peopleSearchStatus");
  if(status){
    status.innerHTML=raw
      ? `<b>${ps.length}</b> colaborador${ps.length===1?"":"es"} encontrado${ps.length===1?"":"s"} para <b>${esc(raw)}</b>`
      : `<b>${all.length}</b> colaboradores registrados`;
  }

  $("#peopleCards").innerHTML=ps.map(p=>{
    const grouped=groupCount(p.items,"equipo");
    const top=topEntries(grouped,3).map(([k,v])=>`${k} (${v})`).join(" · ");
    return `
      <article class="person-card person-card-search" data-person="${esc(p.key)}" tabindex="0" role="button" aria-label="Ver equipos de ${esc(p.name)}">
        <div class="person-card-top">
          <div>
            <h4>${esc(p.name)}</h4>
            <div class="person-dni">DNI: <b>${esc(p.dni||"—")}</b></div>
          </div>
          <div class="person-equipment-count">${p.items.length}</div>
        </div>
        <div class="card-meta">${p.items.length===1?"equipo / registro a cargo":"equipos / registros a cargo"}<br>${esc([...p.sites].join(", ")||"Sin sede registrada")}</div>
        ${top?`<div class="person-equipment-preview"><b>Utiliza:</b> ${esc(top)}</div>`:"<div class=\"person-empty-tag\">Sin equipos asignados</div>"}
        <button class="person-view-btn" type="button" data-person-open="${esc(p.key)}">Ver equipos que utiliza</button>
        ${canEdit()?`<div class="person-manage-actions">
          <button class="btn btn-soft btn-small" type="button" data-person-edit="${esc(p.key)}">Editar colaborador</button>
          <button class="btn btn-danger btn-small" type="button" data-person-remove="${esc(p.key)}">Quitar</button>
        </div>`:""}
      </article>`;
  }).join("") || `
    <div class="people-no-results">
      <b>No se encontró ningún colaborador.</b>
      <span>Revisa el nombre o DNI. Los colaboradores creados manualmente también aparecen en esta búsqueda.</span>
    </div>`;

  const byKey=key=>collaborators().find(x=>x.key===key);
  const openByKey=key=>{const p=byKey(key);if(p)openPerson(p.key)};
  $$("[data-person]").forEach(c=>{
    c.onclick=e=>{
      if(e.target.closest("[data-person-open],[data-person-edit],[data-person-remove]")) return;
      openByKey(c.dataset.person);
    };
    c.onkeydown=e=>{
      if((e.key==="Enter"||e.key===" ")&&!e.target.closest("button")){e.preventDefault();openByKey(c.dataset.person)}
    };
  });
  $$("[data-person-open]").forEach(b=>b.onclick=e=>{e.stopPropagation();openByKey(b.dataset.personOpen)});
  $$("[data-person-edit]").forEach(b=>b.onclick=e=>{e.stopPropagation();const p=byKey(b.dataset.personEdit);if(p)openEditCollaborator(p)});
  $$("[data-person-remove]").forEach(b=>b.onclick=e=>{e.stopPropagation();const p=byKey(b.dataset.personRemove);if(p)openRemoveCollaborator(p)});
}
function openCollaboratorAssignment(prefillKey="") {
  if(!ensureEditable()) return;

  const people=collaborators();
  const isWarehouseAsset=r=>{
    const area=norm(r?.area||"");
    return (r?.locationType==="ALMACEN" || area.includes("ALMAC")) && !String(r?.responsable||"").trim();
  };
  const warehouseCount=activeInventory().filter(isWarehouseAsset).length;
  const destinationOptions=people.map(p=>`<option value="${esc(p.key)}">${esc(p.name)}${p.dni?` · DNI ${esc(p.dni)}`:""}</option>`).join("");

  showModal("Colaboradores · asignación de equipos","Registra un colaborador nuevo o transfiere varios equipos desde Almacén TIC u otro colaborador.",`
    <div id="collabAssignAlert"></div>

    <div class="assignment-summary">
      <div><b>Asignación múltiple</b><span>Agrega los componentes que quedarán a nombre del colaborador. La selección se mantendrá visible arriba hasta que guardes.</span></div>
      <div class="assignment-summary-number"><b>${warehouseCount}</b><span>en almacén</span></div>
      <div class="assignment-summary-number"><b id="caSelectedCount">0</b><span>agregados</span></div>
    </div>

    <div class="form-grid collaborator-form">
      <div class="form-field full">
        <label>Colaborador destino *</label>
        <select id="caDestination">
          <option value="__new__">+ Registrar nuevo colaborador</option>
          ${destinationOptions}
        </select>
      </div>

      <div class="form-field full"><label>Nombres y apellidos *</label><input id="caName" placeholder="Ej. PÉREZ GARCÍA, JUAN CARLOS"></div>
      <div class="form-field"><label>DNI *</label><input id="caDni" maxlength="8" inputmode="numeric" placeholder="8 dígitos"></div>
      <div class="form-field"><label>Sede *</label><input id="caSite" placeholder="Ej. CENTRO"></div>
      <div class="form-field full"><label>Área / aula donde utilizará los equipos *</label><input id="caArea" placeholder="Ej. SECRETARÍA / AULA / COORDINACIÓN"></div>
    </div>

    <section class="selected-assets-panel">
      <div class="selected-assets-head">
        <div>
          <b>Equipos agregados para este colaborador</b>
          <span>Lo que agregues quedará aquí aunque cambies el buscador o el origen. Puedes editar el destino por equipo o quitarlo antes de guardar.</span>
        </div>
        <button id="caClearSelection" class="btn btn-soft" type="button">Quitar todos</button>
      </div>
      <div id="caSelectedAssets"></div>
    </section>

    <div class="assignment-divider">
      <div><b>Buscar equipos para agregar</b><span>Puedes traer equipos del almacén o transferir equipos actualmente asignados a otro colaborador.</span></div>
    </div>

    <div class="assignment-filters">
      <div class="form-field">
        <label>Origen de los equipos</label>
        <select id="caSource">
          <option value="ALMACEN">Almacén TIC</option>
          <option value="COLABORADOR">Otro colaborador</option>
          <option value="TODOS">Almacén + colaboradores</option>
        </select>
      </div>
      <div class="form-field" id="caOriginWrap">
        <label>Colaborador actual</label>
        <select id="caOrigin"></select>
      </div>
      <div class="form-field assignment-search-field">
        <label>Buscar componente</label>
        <input id="caAssetSearch" placeholder="Código TIC, equipo, marca, serie...">
      </div>
    </div>

    <div class="asset-picker-toolbar">
      <div class="button-row">
        <button id="caAddVisible" class="btn btn-soft" type="button">Agregar visibles</button>
      </div>
      <span id="caVisibleCount" class="muted"></span>
    </div>

    <div id="caAssetPicker"></div>

    <div class="form-field full assignment-note">
      <label>Motivo / observación general</label>
      <textarea id="caObservation" placeholder="Ej. entrega de equipos, cambio de responsable, renovación de puesto...">Asignación de equipos al colaborador</textarea>
    </div>

    <div class="modal-actions assignment-actions">
      <button id="saveCollaboratorAssignment" class="btn btn-primary">Guardar colaborador</button>
    </div>
  `);

  const selected=new Map();
  let editingSelectedId="";
  const destination=$("#caDestination");
  const source=$("#caSource");
  const origin=$("#caOrigin");
  const originWrap=$("#caOriginWrap");
  const search=$("#caAssetSearch");
  const nameInput=$("#caName");
  const dniInput=$("#caDni");
  const siteInput=$("#caSite");
  const areaInput=$("#caArea");

  const samePerson=(r,p)=>{
    if(!p || !r?.responsable) return false;
    const rdni=String(r.dni||"").replace(/\D/g,"");
    const pdni=String(p.dni||"").replace(/\D/g,"");
    return Boolean(rdni && pdni && rdni===pdni) || norm(r.responsable)===norm(p.name);
  };

  const getDestination=()=>destination.value==="__new__" ? null : people.find(p=>p.key===destination.value);

  const updateSelectedCount=()=>{
    if($("#caSelectedCount")) $("#caSelectedCount").textContent=String(selected.size);
    const save=$("#saveCollaboratorAssignment");
    if(save) save.textContent=selected.size ? `Guardar colaborador y asignar ${selected.size} equipo${selected.size===1?"":"s"}` : "Guardar colaborador";
  };

  const selectedDestinationLabel=(meta)=>{
    const sede=String(meta?.sede||"").trim() || siteInput.value.trim().toUpperCase();
    const area=String(meta?.area||"").trim() || areaInput.value.trim();
    return [sede,area].filter(Boolean).join(" / ") || "Se definirá con la sede y área del colaborador";
  };

  const renderSelectedAssets=()=>{
    updateSelectedCount();
    const host=$("#caSelectedAssets");
    if(!host) return;

    const rows=[...selected.entries()]
      .map(([id,meta])=>({r:activeInventory().find(x=>x.id===id),meta}))
      .filter(x=>x.r);

    if(!rows.length){
      host.innerHTML=`<div class="selected-assets-empty">Aún no has agregado equipos. Busca un componente abajo y pulsa <b>Agregar</b>.</div>`;
      return;
    }

    host.innerHTML=`
      <div class="table-wrap selected-assets-table">
        <table>
          <thead><tr><th>Código TIC</th><th>Equipo / componente</th><th>Origen actual</th><th>Destino</th><th>Acciones</th></tr></thead>
          <tbody>${rows.map(({r,meta})=>`
            <tr class="selected-confirmed-row">
              <td><span class="code">${esc(r.codigo||r.id)}</span></td>
              <td><b>${esc(r.equipo||"Sin tipo")}</b><div class="muted">${esc([r.marca,r.modelo,r.serie].filter(Boolean).join(" / "))}</div></td>
              <td>${esc([r.sede,r.area,r.responsable||"ALMACÉN TIC"].filter(Boolean).join(" / "))}</td>
              <td>${esc(selectedDestinationLabel(meta))}${(meta.sede||meta.area||meta.observation)?'<div class="selected-custom-tag">Destino personalizado</div>':""}</td>
              <td>
                <div class="selected-actions">
                  <button class="btn btn-soft btn-small" type="button" data-edit-selected="${esc(r.id)}">Editar</button>
                  <button class="btn btn-danger btn-small" type="button" data-remove-selected="${esc(r.id)}">Quitar</button>
                </div>
              </td>
            </tr>
            ${editingSelectedId===r.id?`
            <tr class="selected-edit-row">
              <td colspan="5">
                <div class="selected-edit-box">
                  <div class="form-grid">
                    <div class="form-field"><label>Sede para este equipo</label><input id="caEditSite" value="${esc(meta.sede||"")}" placeholder="${esc(siteInput.value.trim().toUpperCase()||"Usar sede general")}"></div>
                    <div class="form-field"><label>Área / aula para este equipo</label><input id="caEditArea" value="${esc(meta.area||"")}" placeholder="${esc(areaInput.value.trim()||"Usar área general")}"></div>
                    <div class="form-field full"><label>Observación específica</label><textarea id="caEditObservation" placeholder="Opcional. Si queda vacío se usará la observación general.">${esc(meta.observation||"")}</textarea></div>
                  </div>
                  <div class="modal-actions">
                    <button class="btn btn-primary btn-small" type="button" data-save-selected="${esc(r.id)}">Guardar edición</button>
                    <button class="btn btn-soft btn-small" type="button" data-cancel-selected="${esc(r.id)}">Cancelar</button>
                  </div>
                </div>
              </td>
            </tr>`:""}
          `).join("")}
          </tbody>
        </table>
      </div>`;

    $$("[data-remove-selected]").forEach(btn=>btn.onclick=()=>{
      selected.delete(btn.dataset.removeSelected);
      if(editingSelectedId===btn.dataset.removeSelected) editingSelectedId="";
      renderSelectedAssets();
      renderAssetPicker();
    });

    $$("[data-edit-selected]").forEach(btn=>btn.onclick=()=>{
      editingSelectedId=btn.dataset.editSelected;
      renderSelectedAssets();
    });

    $$("[data-cancel-selected]").forEach(btn=>btn.onclick=()=>{
      editingSelectedId="";
      renderSelectedAssets();
    });

    $$("[data-save-selected]").forEach(btn=>btn.onclick=()=>{
      const id=btn.dataset.saveSelected;
      const meta=selected.get(id)||{};
      meta.sede=$("#caEditSite")?.value.trim().toUpperCase()||"";
      meta.area=$("#caEditArea")?.value.trim()||"";
      meta.observation=$("#caEditObservation")?.value.trim()||"";
      selected.set(id,meta);
      editingSelectedId="";
      renderSelectedAssets();
    });
  };

  const syncOriginOptions=()=>{
    const dest=getDestination();
    const current=origin.value;
    const available=people.filter(p=>p.items.length>0 && (!dest || p.key!==dest.key));
    origin.innerHTML=available.length
      ? available.map(p=>`<option value="${esc(p.key)}">${esc(p.name)} · ${p.items.length} equipo(s)</option>`).join("")
      : `<option value="">Sin colaboradores con equipos</option>`;
    if(available.some(p=>p.key===current)) origin.value=current;
    originWrap.classList.toggle("hidden",source.value!=="COLABORADOR");
  };

  const assetRows=()=>{
    const q=norm(search.value||"");
    const dest=getDestination();
    const originPerson=people.find(p=>p.key===origin.value);

    return activeInventory().filter(r=>{
      if(dest && samePerson(r,dest)) return false;

      let matchesSource=false;
      if(source.value==="ALMACEN") matchesSource=isWarehouseAsset(r);
      else if(source.value==="COLABORADOR") matchesSource=Boolean(originPerson && samePerson(r,originPerson));
      else matchesSource=isWarehouseAsset(r) || Boolean(r.responsable);

      if(!matchesSource) return false;
      if(!q) return true;
      return norm([r.codigo,r.id,r.equipo,r.descripcion,r.marca,r.modelo,r.serie,r.sede,r.area,r.responsable,r.estado].join(" ")).includes(q);
    }).sort((a,b)=>{
      const ta=`${a.equipo||""} ${a.codigo||a.id||""}`;
      const tb=`${b.equipo||""} ${b.codigo||b.id||""}`;
      return ta.localeCompare(tb,"es");
    });
  };

  const addAsset=id=>{
    if(!id || selected.has(id)) return;
    const r=activeInventory().find(x=>x.id===id);
    if(!r) return;
    selected.set(id,{sede:"",area:"",observation:""});
    renderSelectedAssets();
    renderAssetPicker();
  };

  const renderAssetPicker=()=>{
    syncOriginOptions();
    const rows=assetRows();
    if($("#caVisibleCount")) $("#caVisibleCount").textContent=`${rows.length} equipo(s) disponibles con este filtro`;

    $("#caAssetPicker").innerHTML=rows.length ? `
      <div class="table-wrap asset-picker-table">
        <table>
          <thead><tr><th>Código TIC</th><th>Equipo / componente</th><th>Marca / modelo</th><th>Serie</th><th>Ubicación actual</th><th>Responsable actual</th><th>Estado</th><th></th></tr></thead>
          <tbody>${rows.map(r=>{
            const added=selected.has(r.id);
            return `
            <tr class="${added?"asset-selected":""}">
              <td><span class="code">${esc(r.codigo||r.id)}</span></td>
              <td><b>${esc(r.equipo||"Sin tipo")}</b><div class="muted">${esc(r.descripcion||"")}</div></td>
              <td>${esc([r.marca,r.modelo].filter(Boolean).join(" / ")||"—")}</td>
              <td>${esc(r.serie||"—")}</td>
              <td>${esc([r.sede,r.area].filter(Boolean).join(" / ")||"—")}</td>
              <td>${esc(r.responsable||"ALMACÉN TIC")}</td>
              <td><span class="badge ${statusClass(r.estado)}">${esc(r.estado||"SIN ESTADO")}</span></td>
              <td><button class="btn ${added?"btn-success":"btn-primary"} btn-small" type="button" data-add-asset="${esc(r.id)}" ${added?"disabled":""}>${added?"Agregado":"Agregar"}</button></td>
            </tr>`;
          }).join("")}
          </tbody>
        </table>
      </div>` : `<div class="asset-picker-empty">No hay equipos disponibles para el origen y filtro seleccionado.</div>`;

    $$("[data-add-asset]").forEach(btn=>btn.onclick=()=>addAsset(btn.dataset.addAsset));
  };

  const syncDestinationFields=()=>{
    const p=getDestination();

    if(p){
      nameInput.value=p.name||"";
      dniInput.value=p.dni||"";
      nameInput.readOnly=true;
      dniInput.readOnly=true;
      siteInput.value=[...p.sites][0]||"";
      areaInput.value=[...p.areas][0]||"";
    }else{
      nameInput.readOnly=false;
      dniInput.readOnly=false;
      nameInput.value="";
      dniInput.value="";
      siteInput.value="";
      areaInput.value="";
    }
    syncOriginOptions();
    renderSelectedAssets();
    renderAssetPicker();
  };

  destination.onchange=syncDestinationFields;
  source.onchange=()=>{syncOriginOptions();renderAssetPicker()};
  origin.onchange=renderAssetPicker;
  search.oninput=renderAssetPicker;
  siteInput.oninput=renderSelectedAssets;
  areaInput.oninput=renderSelectedAssets;

  $("#caAddVisible").onclick=()=>{
    assetRows().forEach(r=>{
      if(!selected.has(r.id)) selected.set(r.id,{sede:"",area:"",observation:""});
    });
    renderSelectedAssets();
    renderAssetPicker();
  };

  $("#caClearSelection").onclick=()=>{
    if(!selected.size) return;
    if(!confirm(`¿Quitar los ${selected.size} equipo(s) agregados de esta asignación?`)) return;
    selected.clear();
    editingSelectedId="";
    renderSelectedAssets();
    renderAssetPicker();
  };

  if(prefillKey && people.some(p=>p.key===prefillKey)){
    destination.value=prefillKey;
  }
  syncDestinationFields();

  $("#saveCollaboratorAssignment").onclick=async()=>{
    const alert=$("#collabAssignAlert");
    const isNew=destination.value==="__new__";
    const name=nameInput.value.trim();
    const dni=dniInput.value.replace(/\D/g,"").slice(0,8);
    const defaultSite=siteInput.value.trim().toUpperCase();
    const defaultArea=areaInput.value.trim();
    const generalObservation=$("#caObservation").value.trim()||"Asignación de equipos al colaborador";
    const records=[...selected.entries()]
      .map(([id,meta])=>({r:activeInventory().find(r=>r.id===id),meta}))
      .filter(x=>x.r);

    alert.className="";
    alert.textContent="";

    if(!name){
      alert.className="alert error";alert.textContent="Ingresa los nombres y apellidos del colaborador.";return;
    }
    if(isNew && !/^\d{8}$/.test(dni)){
      alert.className="alert error";alert.textContent="Para un colaborador nuevo, el DNI debe tener 8 dígitos.";return;
    }
    if(dni && !/^\d{8}$/.test(dni)){
      alert.className="alert error";alert.textContent="El DNI debe tener 8 dígitos.";return;
    }
    if(!defaultSite || !defaultArea){
      alert.className="alert error";alert.textContent="Completa la sede y el área donde utilizará los equipos.";return;
    }

    const duplicateDni=(state?.collaboratorDirectory||[]).find(c=>
      isNew && String(c.dni||"").replace(/\D/g,"")===dni && norm(c.name)!==norm(name)
    );
    if(duplicateDni){
      alert.className="alert error";alert.textContent=`El DNI ${dni} ya está registrado a nombre de ${duplicateDni.name}.`;return;
    }

    upsertCollaboratorDirectory({name,dni,sede:defaultSite,area:defaultArea});
    state.webMovements=state.webMovements||[];

    records.forEach(({r,meta})=>{
      const oldSite=r.sede||"";
      const oldArea=r.area||"";
      const oldResp=r.responsable||"";
      const oldDni=r.dni||"";
      const from=[oldSite,oldArea,oldResp].filter(Boolean).join(" / ") || "SIN UBICACIÓN";
      const newSite=String(meta?.sede||"").trim().toUpperCase() || defaultSite;
      const newArea=String(meta?.area||"").trim() || defaultArea;
      const observation=String(meta?.observation||"").trim() || generalObservation;

      if(oldResp && norm(oldResp)!==norm(name)){
        upsertCollaboratorDirectory({name:oldResp,dni:oldDni,sede:oldSite,area:oldArea});
      }

      r.history=r.history||[];
      r.history.unshift({
        type:"WEB",
        source:"Asignación de colaborador",
        fecha:today(),
        sede:oldSite,
        area:oldArea,
        responsable:oldResp,
        estado:r.estado,
        observaciones:`Ubicación anterior. ${observation}`
      });

      r.sede=newSite;
      r.area=newArea;
      r.responsable=name;
      r.dni=dni;
      r.locationType="ASIGNADO";
      r.needsReview=false;
      r.conflictLocations=[];
      r.conflictRecords=[];

      state.webMovements.unshift({
        id:crypto.randomUUID(),
        fecha:today(),
        source:oldResp ? "Transferencia de colaborador" : "Asignación desde almacén",
        codigo:r.codigo,
        equipo:r.equipo,
        from,
        to:[newSite,newArea,name].filter(Boolean).join(" / "),
        responsable:name,
        motivo:observation,
        observaciones:observation
      });
    });

    try{
      await persist(records.length ? "ASIGNAR_EQUIPOS_COLABORADOR" : "ALTA_COLABORADOR");
    }catch(_){
      return;
    }

    closeModal();
    renderAll();

    if(records.length){
      toast(`${name}: ${records.length} equipo(s) asignado(s) correctamente.`,"success");
    }else{
      toast(`Colaborador ${name} registrado. Puedes asignarle equipos cuando lo necesites.`,"success");
    }
  };
}
function openPerson(key){
  const p=collaborators().find(x=>x.key===key);if(!p)return;
  const items=p.items.slice().sort((a,b)=>(a.equipo||"").localeCompare(b.equipo||"","es") || (a.codigo||a.id||"").localeCompare(b.codigo||b.id||"","es"));
  const grouped=topEntries(groupCount(items,"equipo"),10);
  const latestCargo=latestPrintedCargoForPerson(p);
  const summary=grouped.length
    ? `<div class="person-usage-summary">${grouped.map(([name,count])=>`<div><b>${count}</b><span>${esc(name)}</span></div>`).join("")}</div>`
    : `<div class="asset-picker-empty">Este colaborador aún no tiene equipos asignados.</div>`;
  const cargoStatus=latestCargo
    ? `<div class="cargo-status-box">
        <div><span>Cargo</span><b>${esc(latestCargo.code)}</b></div>
        <div><span>Impreso</span><b>${esc(fmtDateTime(latestCargo.printedAt))}</b></div>
        <div><span>Recepción</span><b class="${latestCargo.receivedSigned?"cargo-received":"cargo-pending"}">${latestCargo.receivedSigned?"FIRMADO / RECIBIDO":"PENDIENTE DE FIRMA"}</b></div>
      </div>`
    : `<div class="alert info">Este colaborador todavía no tiene un cargo impreso.</div>`;

  showModal(p.name,`${items.length} equipos / registros actualmente a cargo`,`
    <div class="detail-grid">
      <div class="detail-field"><span>DNI</span><b>${esc(p.dni||"—")}</b></div>
      <div class="detail-field"><span>Sede(s)</span><b>${esc([...p.sites].join(", ")||"—")}</b></div>
      <div class="detail-field"><span>Área(s)</span><b>${esc([...p.areas].join(", ")||"—")}</b></div>
    </div>
    <h4 class="section-title">Actualmente utiliza</h4>
    ${summary}
    <h4 class="section-title">Control del cargo</h4>
    ${cargoStatus}
    <div class="modal-actions">
      ${canEdit()?'<button id="assignPersonAssets" class="btn btn-primary">Asignar / transferir equipos</button>':""}
      ${canEdit()?'<button id="editCollaboratorProfile" class="btn btn-soft">Editar colaborador</button>':""}
      <button id="editPersonCargo" class="btn btn-soft">Editar todo el cargo</button>
      <button id="printCargo" class="btn btn-primary">Imprimir cargo</button>
      ${latestCargo&&canEdit()?`<button id="toggleCargoReceived" class="btn ${latestCargo.receivedSigned?"btn-warning":"btn-success"}">${latestCargo.receivedSigned?"Quitar recibido":"Marcar firmado y recibido"}</button>`:""}
      ${canEdit()?'<button id="removeCollaboratorProfile" class="btn btn-danger">Quitar colaborador</button>':""}
    </div>
    <h4 class="section-title">Detalle de equipos a cargo</h4>
    ${items.length?tableHtml(items,false):'<div class="asset-picker-empty">Sin equipos asignados.</div>'}
  `);

  if($("#assignPersonAssets")) $("#assignPersonAssets").onclick=()=>openCollaboratorAssignment(p.key);
  if($("#editCollaboratorProfile")) $("#editCollaboratorProfile").onclick=()=>openEditCollaborator(p);
  if($("#removeCollaboratorProfile")) $("#removeCollaboratorProfile").onclick=()=>openRemoveCollaborator(p);
  $("#editPersonCargo").onclick=()=>openEditPersonCargo(p);
  $("#printCargo").onclick=()=>printCargo(p);
  if($("#toggleCargoReceived")) $("#toggleCargoReceived").onclick=async()=>{
    const next=!latestCargo.receivedSigned;
    if(!confirm(next?"¿Confirmar que el usuario ya firmó el cargo y recibió los equipos?":"¿Quitar la confirmación de firma y recepción?")) return;
    await setCargoReceived(latestCargo.id,next);
    closeModal();
    openPerson(key);
  };
  wireTables();
}

  function wireAppControls() {
    $$(".nav-item").forEach(b=>b.onclick=()=>setView(b.dataset.view));
    $$("[data-go]").forEach(b=>b.onclick=()=>setView(b.dataset.go));
    if($("#mobileMenu")) $("#mobileMenu").onclick=()=>$("#sidebar")?.classList.toggle("open");
    if($("#modalClose")) $("#modalClose").onclick=closeModal;
    // Política global de modales: nunca cerrar al hacer clic fuera.
    // Aplica a fichas, altas, edición, movimientos, colaboradores, usuarios,
    // revisión y cualquier ventana que utilice #modal.
    if($("#modal")){
      $("#modal").onclick=(e)=>{
        if(e.target === $("#modal")){
          e.preventDefault();
          e.stopPropagation();
        }
      };
    }

    // Evita cierres accidentales con Escape. El cierre manual se realiza
    // únicamente con la X; los botones Guardar/Confirmar pueden cerrar
    // después de completar correctamente su operación.
    document.addEventListener("keydown",(e)=>{
      if(e.key==="Escape" && $("#modal")?.classList.contains("open")){
        e.preventDefault();
        e.stopPropagation();
      }
    });
    if($("#dismissPrivacy")) $("#dismissPrivacy").onclick=()=>$("#privacyBanner")?.remove();
    if($("#importBtnTop")) $("#importBtnTop").onclick=()=>$("#excelInput")?.click();
    if($("#importBtnEmpty")) $("#importBtnEmpty").onclick=()=>$("#excelInput")?.click();
    if($("#excelInput")) $("#excelInput").onchange=e=>{const file=e.target.files?.[0];if(file)importExcel(file);e.target.value=""};
    if($("#demoInfoBtn")) $("#demoInfoBtn").onclick=demoInfo;
    if($("#addAssetBtn")) $("#addAssetBtn").onclick=()=>openAddAsset();
    if($("#manageCollaboratorBtn")) $("#manageCollaboratorBtn").onclick=()=>openCollaboratorAssignment();
    if($("#exportCsvBtn")) $("#exportCsvBtn").onclick=exportInventory;
    if($("#backupBtn")) $("#backupBtn").onclick=downloadBackup;
    if($("#exportMovementsBtn")) $("#exportMovementsBtn").onclick=exportMovements;
    if($("#restoreBtn")) $("#restoreBtn").onclick=()=>$("#restoreInput")?.click();
    if($("#restoreInput")) $("#restoreInput").onchange=e=>{const file=e.target.files?.[0];if(file)restoreBackup(file);e.target.value=""};
    if($("#clearLocalBtn")) $("#clearLocalBtn").onclick=clearLocalData;
    if($("#createUserBtn")) $("#createUserBtn").onclick=openCreateUser;
    if($("#refreshUsersBtn")) $("#refreshUsersBtn").onclick=renderUsers;
    if($("#refreshAuditBtn")) $("#refreshAuditBtn").onclick=()=>{renderCargoAudit();renderAudit()};
    if($("#refreshCargoAuditBtn")) $("#refreshCargoAuditBtn").onclick=renderCargoAudit;
    if($("#cargoAuditSearch")) $("#cargoAuditSearch").addEventListener("input",renderCargoAudit);
    if($("#cargoAuditSite")) $("#cargoAuditSite").addEventListener("change",renderCargoAudit);
    if($("#cargoAuditStatus")) $("#cargoAuditStatus").addEventListener("change",renderCargoAudit);
    if($("#authorizedSearch")) $("#authorizedSearch").addEventListener("input",renderUsers);
    if($("#authorizedStatus")) $("#authorizedStatus").addEventListener("change",renderUsers);
    if($("#authorizedRole")) $("#authorizedRole").addEventListener("change",renderUsers);

    const doLogout=async()=>{
      if(realtimeChannel){await PraxisCloud.unsubscribe(realtimeChannel);realtimeChannel=null;}
      if(accessChannel){await PraxisCloud.unsubscribe(accessChannel);accessChannel=null;}
      await PraxisCloud.signOut();
      state=null;currentProfile=null;
      if($("#loginPassword")) $("#loginPassword").value="";
      showLogin();
    };
    if($("#logoutBtn")) $("#logoutBtn").onclick=doLogout;
    if($("#logoutSidebarBtn")) $("#logoutSidebarBtn").onclick=doLogout;

    ["#inventorySearch","#siteFilter","#categoryFilter","#statusFilter","#locationFilter"].forEach(sel=>{
      const el=$(sel); if(el) el.addEventListener("input",()=>{page=1;renderInventory()});
    });
    if($("#peopleSearch")) $("#peopleSearch").addEventListener("input",renderPeople);
    if($("#reviewSearch")) $("#reviewSearch").addEventListener("input",renderReview);
    if($("#globalSearch")) $("#globalSearch").addEventListener("input",e=>{
      if(!state?.inventory?.length)return;
      const value=e.target.value;

      if(currentView==="people"){
        if($("#peopleSearch")) $("#peopleSearch").value=value;
        renderPeople();
        return;
      }

      if($("#inventorySearch")) $("#inventorySearch").value=value;
      page=1;
      if(value.trim())setView("inventory");
      renderInventory();
    });
  }

  async function init() {
    cloudMode = PraxisCloud.configured();

    wireAuthControls();
    wireAppControls();

    if (cloudMode) {
      showLogin();
      if ($("#loginPassword")) $("#loginPassword").value = "";
      setTimeout(() => $("#loginEmail")?.focus(), 50);
    } else {
      state = await PraxisDB.get(STATE_KEY);
      $("#authGate")?.classList.add("hidden");
      $("#appShell")?.classList.remove("hidden");
      updateProfileUI();
      renderAll();
    }
  }

  document.addEventListener("DOMContentLoaded",()=>{
    init().catch(err=>{
      console.error("Inventario Praxis no pudo iniciar",err);
      try{
        const gate=document.querySelector("#authGate");
        const shell=document.querySelector("#appShell");
        const box=document.querySelector("#authMessage");
        if(shell) shell.classList.add("hidden");
        if(gate) gate.classList.remove("hidden");
        if(box){
          box.textContent="No se pudo iniciar el sistema. Actualiza la página. Si continúa, comunícate con TIC.";
          box.className="auth-message show error";
        }
      }catch(_){}
    });
  });
})();
