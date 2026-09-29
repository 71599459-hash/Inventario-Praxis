
(() => {
  const STATE_KEY = "praxis-state-v2";
  const PAGE_SIZE = 30;
  let state = null;
  let currentView = "dashboard";
  let page = 1;
  let selectedId = null;
  let cloudMode = false;
  let currentProfile = null;
  let realtimeChannel = null;

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
    await PraxisDB.set(STATE_KEY, state);
    if (!cloudMode || !PraxisCloud.configured()) return true;
    try {
      await PraxisCloud.saveState(state, action);
      return true;
    } catch (err) {
      console.error(err);
      if (err && err.message === "VERSION_CONFLICT") {
        const latest = await PraxisCloud.loadState();
        state = latest.state;
        await PraxisDB.set(STATE_KEY, state);
        renderAll();
        toast("Otro usuario actualizó el inventario. Se recargó la versión más reciente; vuelve a realizar tu cambio.","error");
        throw err;
      }
      toast((err && err.message) || "No se pudo guardar en la base de datos.","error");
      throw err;
    }
  }

  function canEdit() {
    return !cloudMode || PraxisCloud.canEdit();
  }

  function isAdmin() {
    return cloudMode && PraxisCloud.isAdmin();
  }

  function ensureEditable() {
    if (canEdit()) return true;
    toast("Tu usuario tiene rol CONSULTA. No puede modificar el inventario.","error");
    return false;
  }

  const roleLabel = role => ({
    ADMIN_TIC:"Administrador TIC",
    EDITOR:"Editor",
    CONSULTA:"Consulta"
  }[role] || role || "Local");

  const fmtDateTime = value => {
    if(!value) return "Nunca";
    const d=new Date(value);
    return Number.isNaN(d.getTime()) ? String(value) : d.toLocaleString("es-PE",{dateStyle:"short",timeStyle:"short"});
  };

  function updateProfileUI() {
    const p = currentProfile || PraxisCloud.getProfile();
    const displayName=(p && (p.nombre || p.email)) || "Usuario";
    const displayRole=roleLabel((p && p.role) || "LOCAL");
    if ($("#profileName")) $("#profileName").textContent = displayName;
    if ($("#profileRole")) $("#profileRole").textContent = displayRole;
    if ($("#sidebarProfileName")) $("#sidebarProfileName").textContent = displayName;
    if ($("#sidebarProfileRole")) $("#sidebarProfileRole").textContent = displayRole;
    $$(".edit-only").forEach(el=>el.classList.toggle("hidden",cloudMode && !canEdit()));
    $$(".admin-only").forEach(el=>el.classList.toggle("hidden",!isAdmin()));
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
      people:["Colaboradores","Equipos y componentes a cargo"],
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
    if (view === "users" && isAdmin()) { renderUsers(); renderAudit(); }
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
      r.sede=newSite;r.area=newArea;r.responsable=newResp;r.locationType=norm(newArea).includes("ALMAC")?"ALMACEN":newResp?"ASIGNADO":"SEDE";
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
        <button id="printCargo" class="btn btn-primary">Imprimir cargo</button>
        <button id="filterPerson" class="btn btn-soft">Ver en inventario</button>
      </div>
      <div class="alert info"><b>Ficha técnica de cargo:</b> se imprime una sola hoja con la <b>Revisión Inicial</b> fechada y la <b>Revisión Final</b> en blanco. Esa misma hoja física se conserva y se completa al finalizar la entrega.</div>
      <h4 class="section-title">Equipos a cargo</h4>${tableHtml(items,false)}
    `);
    $("#printCargo").onclick=()=>printCargo(p);
    $("#filterPerson").onclick=()=>{closeModal();setView("inventory");$("#inventorySearch").value=p.name;page=1;renderInventory()};
    wireTables();
  }

  const TIC_HEAD_NAME="FABIÁN PUENTE, FRANK JAIME";
  const CARGO_BASE_SEQUENCE=16;

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
    const existing=(state?.cargoDocuments||[])
      .map(d=>String(d.code||"").match(new RegExp("^FTEC-"+year+"-(\\d+)$","i")))
      .filter(Boolean)
      .map(m=>Number(m[1]))
      .filter(Number.isFinite);
    const seq=Math.max(CARGO_BASE_SEQUENCE,...existing)+1;
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
    const snap=cargoSnapshot(p.items);
    const fingerprint=cargoFingerprint(snap);
    const currentDate=today();

    // Reimpresión segura: si hoy ya se generó exactamente el mismo cargo,
    // reutiliza el mismo código FTEC en lugar de crear otro.
    let doc=[...state.cargoDocuments].reverse().find(d=>
      d.personKey===key &&
      d.initialDate===currentDate &&
      cargoFingerprint(d.items||[])===fingerprint
    );

    if(!doc){
      doc={
        id:crypto.randomUUID(),
        code:nextCargoCode(),
        personKey:key,
        personName:p.name,
        dni:p.dni||"",
        sede:[...p.sites].join(", "),
        area:[...p.areas].join(", "),
        situacion:"INTERNO",
        initialDate:currentDate,
        finalDate:"",
        itemFingerprint:fingerprint,
        items:snap,
        createdBy:currentProfile?.id||"",
        createdByName:currentProfile?.nombre||currentProfile?.email||"",
        createdAt:new Date().toISOString()
      };
      state.cargoDocuments.push(doc);
      await persist("IMPRIMIR_CARGO");
    }

    const technician=(currentProfile?.nombre||"CÁRDENAS CURISINCHE, ROBERTO ALEJANDRO").toUpperCase();
    const items=doc.items||[];
    const pageSize=25;
    const pages=[];
    for(let i=0;i<Math.max(1,Math.ceil(items.length/pageSize));i++){
      pages.push(items.slice(i*pageSize,(i+1)*pageSize));
    }

    const logo=document.querySelector(".brand img")?.src||document.querySelector("[data-brand-logo]")?.src||"";
    const pageHtml=pages.map((pageItems,pageIndex)=>{
      const globalStart=pageIndex*pageSize;
      const rows=[];

      for(let i=0;i<pageSize;i++){
        const r=pageItems[i];
        if(r){
          const brandModel=[r.marca,r.modelo].filter(Boolean).join(" / ");
          const codeSerie=[r.codigo,r.serie].filter(Boolean).join(" / ");
          rows.push(`<tr>
            <td class="num">${globalStart+i+1}</td>
            <td class="num">${esc(r.cantidad||"1")}</td>
            <td>${esc(r.equipo)}</td>
            <td>${esc(cargoItemType(r))}</td>
            <td>${esc(brandModel)}</td>
            <td>${esc(r.caracteristicas)}</td>
            <td>${esc(codeSerie||"SIN CÓDIGO")}</td>
            <td>${esc(r.estado)}</td>
            <td class="nowrap">${esc(cargoDate(doc.initialDate))}</td>
            <td class="sign-row"></td>
            <td class="sign-row"></td>
            <td>${esc(r.observaciones)}</td>
          </tr>`);
        }else{
          rows.push(`<tr>
            <td class="num">${globalStart+i+1}</td><td></td><td></td><td></td><td></td><td></td>
            <td></td><td></td><td></td><td class="sign-row"></td><td class="sign-row"></td><td></td>
          </tr>`);
        }
      }

      const isLast=pageIndex===pages.length-1;
      const reviews=isLast?`
        ${cargoSoftwareHtml()}
        <div class="revision-grid">
          ${revisionBox("REVISIÓN INICIAL",doc.initialDate,technician,TIC_HEAD_NAME,"")}
          ${revisionBox("REVISIÓN FINAL","",technician,TIC_HEAD_NAME,"")}
        </div>`:"";

      return `<section class="print-page">
        <div class="doc-title">FICHA TÉCNICA DE EQUIPO DE CÓMPUTO</div>
        <div class="doc-meta">
          <div><b>Código:</b><span>${esc(doc.code)}</span></div>
          <div><b>Situación:</b><span>${esc(doc.situacion||"INTERNO")}</span></div>
          <div class="meta-tech"><b>Técnico:</b><span>${esc(technician)}</span></div>
          <div><b>Sede:</b><span>${esc(doc.sede||"")}</span></div>
          <div><b>Área:</b><span>${esc(doc.area||"")}</span></div>
          <div class="meta-user"><b>Usuario:</b><span>${esc(doc.personName)}${doc.dni?" · DNI: "+esc(doc.dni):""}</span><img src="${logo}"></div>
        </div>

        <div class="section-label">HARDWARE</div>
        <table class="hardware-table">
          <thead><tr>
            <th>Item</th><th>Cant.</th><th>Equipo</th><th>Tipo</th><th>Marca / Modelo</th><th>Características</th>
            <th>Código TIC / Serie</th><th>Estado</th><th>Fecha</th><th>Firma (R.C.)</th><th>Firma TIC</th><th>Observación</th>
          </tr></thead>
          <tbody>${rows.join("")}</tbody>
        </table>

        ${reviews}
        <div class="page-counter">Página ${pageIndex+1} de ${pages.length} · CARGO DE EQUIPOS TIC</div>
      </section>`;
    }).join("");

    const w=window.open("","_blank");
    if(!w){
      toast("El navegador bloqueó la ventana de impresión. Habilita ventanas emergentes para esta página.","error");
      return;
    }

    w.document.write(`<!doctype html><html><head><meta charset="utf-8">
      <title>${esc(doc.code)} - ${esc(doc.personName)}</title>
      <style>
        @page{size:A4 landscape;margin:3mm}
        *{box-sizing:border-box}
        html,body{margin:0;padding:0}
        body{font-family:Arial,Helvetica,sans-serif;color:#111;background:#fff}
        .print-page{width:100%;page-break-after:always;break-after:page;position:relative}
        .print-page:last-child{page-break-after:auto;break-after:auto}

        .doc-title{height:7mm;display:grid;place-items:center;border:1px solid #222;font-weight:800;font-size:9pt;padding:1px}
        .doc-meta{display:grid;grid-template-columns:1fr 1fr 2.1fr;border-left:1px solid #222;border-top:1px solid #222;font-size:6pt}
        .doc-meta>div{display:flex;min-height:5mm;border-right:1px solid #222;border-bottom:1px solid #222;align-items:center}
        .doc-meta b{width:49px;text-align:right;padding-right:3px}
        .doc-meta span{flex:1;padding:1px 3px;font-weight:600}
        .meta-user{position:relative;padding-right:73px}
        .meta-user img{position:absolute;right:4px;top:1px;height:9mm;width:66px;object-fit:contain}

        .section-label{font-size:6pt;font-weight:800;line-height:3.2mm}
        table{width:100%;border-collapse:collapse;table-layout:fixed}
        th,td{border:1px solid #222;padding:.6px 1.5px;vertical-align:middle;overflow-wrap:anywhere;line-height:1.05}
        th{font-size:5.3pt;background:#eee;text-align:center;font-weight:800}
        .hardware-table td{font-size:5.15pt;height:3.75mm}
        .hardware-table th{height:4.2mm}
        .hardware-table th:nth-child(1){width:3%}.hardware-table th:nth-child(2){width:3%}
        .hardware-table th:nth-child(3){width:8%}.hardware-table th:nth-child(4){width:7%}
        .hardware-table th:nth-child(5){width:11%}.hardware-table th:nth-child(6){width:18%}
        .hardware-table th:nth-child(7){width:13%}.hardware-table th:nth-child(8){width:7%}
        .hardware-table th:nth-child(9){width:7%}.hardware-table th:nth-child(10){width:8%}
        .hardware-table th:nth-child(11){width:6%}.hardware-table th:nth-child(12){width:9%}
        .num{text-align:center}.nowrap{white-space:nowrap}.sign-row{height:3.75mm}

        .software-table th{font-size:5pt;height:3.5mm}
        .software-table td{font-size:4.9pt;height:3.25mm}
        .software-table th:nth-child(4n+1){width:3%}
        .software-table th:nth-child(4n+2){width:9%}
        .software-table th:nth-child(4n+3){width:9%}
        .software-table th:nth-child(4n+4){width:12%}

        .revision-grid{display:grid;grid-template-columns:1fr 1fr;gap:8mm;margin-top:2mm;break-inside:avoid}
        .revision-box{border:1px solid #222;font-size:5.2pt;break-inside:avoid}
        .revision-title{text-align:center;font-weight:800;background:#eee;border-bottom:1px solid #222;padding:1px;line-height:3.2mm}
        .revision-date{display:grid;grid-template-columns:40% 60%;border-bottom:1px solid #222;min-height:4mm;align-items:center;text-align:center}
        .revision-date b{border-right:1px solid #222;height:100%;display:grid;place-items:center}
        .revision-observation{min-height:8mm;border-bottom:1px solid #222;text-align:center;display:grid;grid-template-rows:3.2mm 1fr}
        .revision-observation b{background:#eee;border-bottom:1px solid #222}
        .revision-observation span{padding:1px}
        .revision-signatures{display:grid;grid-template-columns:1fr 1fr}
        .signature-cell{text-align:center;border-right:1px solid #222}.signature-cell:last-child{border-right:0}
        .signature-space{height:7mm;border-bottom:1px solid #222}
        .signature-cell b{display:block;padding:1px}
        .revision-names{display:grid;grid-template-columns:1fr 1fr;border-top:1px solid #222}
        .revision-names>div{text-align:center;border-right:1px solid #222}.revision-names>div:last-child{border-right:0}
        .revision-names span{display:block;min-height:3.6mm;padding:1px;border-bottom:1px solid #222}
        .revision-names b{display:block;padding:1px}
        .page-counter{text-align:right;font-size:4.8pt;margin-top:1mm;color:#555}

        @media print{
          body{-webkit-print-color-adjust:exact;print-color-adjust:exact}
        }
      </style>
    </head><body>${pageHtml}
      <script>window.onload=()=>setTimeout(()=>window.print(),250)<\/script>
    </body></html>`);
    w.document.close();
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

  async function renderUsers() {
    if(!isAdmin() || !$("#usersTable")) return;
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

      if($("#authorizedSummary")) $("#authorizedSummary").innerHTML=[
        [allUsers.length,"Cuentas registradas"],
        [activeCount,"Personal autorizado"],
        [pendingCount,"Pendientes / bloqueados"],
        [admins,"Administradores TIC"],
        [editors,"Editores"],
        [consult,"Solo consulta"]
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
        const statusBadge=u.activo
          ? '<span class="badge ok">AUTORIZADO</span>'
          : '<span class="badge warn">PENDIENTE / BLOQUEADO</span>';
        const actionText=u.activo?"Desautorizar":"Autorizar";
        const actionClass=u.activo?"btn-danger":"btn-success";
        return `<tr>
          <td>
            <div class="user-name-cell"><span class="avatar mini">TIC</span><div><b>${esc(u.nombre||"Sin nombre")}</b>${isMe?'<span class="you-tag">Tú</span>':""}<div class="muted">${esc(u.email||"")}</div><div class="muted">DNI: ${esc(u.dni||"—")}</div></div></div>
          </td>
          <td><b>${esc(u.cargo||"—")}</b><div class="muted">${esc([u.sede,u.area].filter(Boolean).join(" / ")||"—")}</div></td>
          <td>
            <select class="role-select" data-role-user="${esc(u.id)}" ${isMe&&u.role==="ADMIN_TIC"&&admins<=1?"disabled":""}>
              <option value="ADMIN_TIC" ${u.role==="ADMIN_TIC"?"selected":""}>Administrador TIC</option>
              <option value="EDITOR" ${u.role==="EDITOR"?"selected":""}>Editor</option>
              <option value="CONSULTA" ${u.role==="CONSULTA"?"selected":""}>Consulta</option>
            </select>
          </td>
          <td>${statusBadge}</td>
          <td>${esc(fmtDateTime(u.last_login_at))}</td>
          <td><button class="btn ${actionClass} btn-small" data-active-user="${esc(u.id)}" data-next-active="${u.activo?"false":"true"}" ${isMe?"disabled":""}>${actionText}</button></td>
        </tr>`;
      }).join("");

      $("#usersTable").innerHTML=`<div class="table-wrap"><table class="authorized-table"><thead><tr><th>Personal</th><th>Cargo / ubicación</th><th>Rol</th><th>Acceso</th><th>Último acceso</th><th>Acción</th></tr></thead><tbody>${rows||'<tr><td colspan="6">No se encontró personal con esos filtros.</td></tr>'}</tbody></table></div>`;

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

      $$("[data-active-user]").forEach(btn=>{
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
      $("#usersTable").innerHTML="<div class='alert error'>"+esc((err&&err.message)||"No se pudo cargar el personal autorizado.")+"</div>";
    }
  }

  async function renderAudit() {
    if(!isAdmin() || !$("#auditTable")) return;
    try {
      const [items,users]=await Promise.all([PraxisCloud.audit(100),PraxisCloud.listUsers()]);
      const userMap=new Map(users.map(u=>[u.id,u]));
      const actionLabel=action=>({
        IMPORTAR_EXCEL:"Importación de Excel",
        ALTA_EQUIPO:"Alta de equipo",
        EDITAR_FICHA:"Edición de ficha",
        MOVIMIENTO_EQUIPO:"Movimiento / transferencia",
        RESOLVER_CONFLICTO:"Resolución de conflicto",
        RESTAURAR_BACKUP:"Restauración de copia",
        IMPRIMIR_CARGO:"Ficha técnica · cargo de equipos",
        CAMBIAR_ROL:"Cambio de rol",
        AUTORIZAR_USUARIO:"Autorización de personal",
        DESAUTORIZAR_USUARIO:"Retiro de acceso",
        BOOTSTRAP_ADMIN:"Alta del primer Administrador TIC",
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
    if(state) await PraxisDB.set(STATE_KEY,state);
    if(realtimeChannel) await PraxisCloud.unsubscribe(realtimeChannel);
    realtimeChannel=PraxisCloud.subscribeState(async latest=>{
      state=latest.state;
      if(state) await PraxisDB.set(STATE_KEY,state);
      renderAll();
      toast("El inventario se actualizó desde otro usuario.","success");
    });
    $("#authGate").classList.add("hidden");
    $("#appShell").classList.remove("hidden");
    updateProfileUI();
    renderAll();
    setView("dashboard");
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
    $("#loginForm").classList.remove("hidden");
    $("#registerForm").classList.add("hidden");
    $("#authSubtitle").textContent="Acceso institucional al inventario centralizado.";
    showAuth(message,type);
  }

  function showRegister(message="",type="info") {
    $("#loginForm").classList.add("hidden");
    $("#registerForm").classList.remove("hidden");
    $("#authSubtitle").textContent="Registro de acceso al sistema.";
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
    const showRegisterBtn=$("#showRegisterBtn");
    if(showRegisterBtn) showRegisterBtn.onclick=()=>{
      const loginEmail=$("#loginEmail");
      const regEmail=$("#regEmail");
      if(regEmail) regEmail.value=PraxisCloud.normalizeInstitutionalEmail(loginEmail?.value||"");
      showRegister();
      setTimeout(()=>$("#regName")?.focus(),50);
    };

    const backBtn=$("#backToLoginBtn");
    if(backBtn) backBtn.onclick=()=>{
      const regEmail=$("#regEmail")?.value?.trim()||"";
      if(regEmail && $("#loginEmail")) $("#loginEmail").value=regEmail.replace(/@praxis\.edu\.pe$/i,"");
      showLogin();
    };

    const registerForm=$("#registerForm");
    if(registerForm) registerForm.onsubmit=async e=>{
      e.preventDefault();
      const nombre=$("#regName")?.value?.trim()||"";
      const dni=$("#regDni")?.value?.trim()||"";
      const cargo=$("#regCargo")?.value?.trim()||"";
      const sede=$("#regSede")?.value?.trim()||"";
      const area=$("#regArea")?.value?.trim()||"";
      const email=$("#regEmail")?.value?.trim()||"";
      const password=$("#regPassword")?.value||"";
      const confirmPassword=$("#regPassword2")?.value||"";

      if(password!==confirmPassword){
        showRegister("Las contraseñas no coinciden.","error");
        return;
      }

      const button=registerForm.querySelector("button[type='submit']");
      const previous=button?.textContent||"Crear acceso al sistema";
      if(button){button.disabled=true;button.textContent="Creando acceso…";}
      try{
        const data=await PraxisCloud.signUp({email,password,nombre,dni,cargo,sede,area});
        if(data?.session){
          await enterCloudApp();
        }else{
          if($("#loginEmail")) $("#loginEmail").value=PraxisCloud.normalizeInstitutionalEmail(email).replace(/@praxis\.edu\.pe$/i,"");
          showLogin("Acceso creado. Ya puedes iniciar sesión.","info");
        }
      }catch(err){
        console.error("Registro falló",err);
        showRegister((err&&err.message)||"No se pudo crear el acceso.","error");
      }finally{
        if(button){button.disabled=false;button.textContent=previous;}
      }
    };

    const loginForm=$("#loginForm");
    if(loginForm) loginForm.onsubmit=async e=>{
      e.preventDefault();
      const email=$("#loginEmail")?.value?.trim()||"";
      const password=$("#loginPassword")?.value||"";
      const box=$("#authMessage");
      if(box){box.textContent="Ingresando…";box.className="auth-message show info";}
      try {
        await PraxisCloud.signIn(email,password);
        await enterCloudApp();
      } catch(err) {
        console.error("Login falló",err);
        showLogin((err&&err.message)||"No se pudo iniciar sesión.","error");
      }
    };
  }

  function wireAppControls() {
    $$(".nav-item").forEach(b=>b.onclick=()=>setView(b.dataset.view));
    $$("[data-go]").forEach(b=>b.onclick=()=>setView(b.dataset.go));
    if($("#mobileMenu")) $("#mobileMenu").onclick=()=>$("#sidebar")?.classList.toggle("open");
    if($("#modalClose")) $("#modalClose").onclick=closeModal;
    if($("#modal")) $("#modal").addEventListener("click",e=>{if(e.target.id==="modal")closeModal()});
    if($("#dismissPrivacy")) $("#dismissPrivacy").onclick=()=>$("#privacyBanner")?.remove();
    if($("#importBtnTop")) $("#importBtnTop").onclick=()=>$("#excelInput")?.click();
    if($("#importBtnEmpty")) $("#importBtnEmpty").onclick=()=>$("#excelInput")?.click();
    if($("#excelInput")) $("#excelInput").onchange=e=>{const file=e.target.files?.[0];if(file)importExcel(file);e.target.value=""};
    if($("#demoInfoBtn")) $("#demoInfoBtn").onclick=demoInfo;
    if($("#addAssetBtn")) $("#addAssetBtn").onclick=openAddAsset;
    if($("#exportCsvBtn")) $("#exportCsvBtn").onclick=exportInventory;
    if($("#backupBtn")) $("#backupBtn").onclick=downloadBackup;
    if($("#exportMovementsBtn")) $("#exportMovementsBtn").onclick=exportMovements;
    if($("#restoreBtn")) $("#restoreBtn").onclick=()=>$("#restoreInput")?.click();
    if($("#restoreInput")) $("#restoreInput").onchange=e=>{const file=e.target.files?.[0];if(file)restoreBackup(file);e.target.value=""};
    if($("#clearLocalBtn")) $("#clearLocalBtn").onclick=clearLocalData;
    if($("#refreshUsersBtn")) $("#refreshUsersBtn").onclick=renderUsers;
    if($("#refreshAuditBtn")) $("#refreshAuditBtn").onclick=renderAudit;
    if($("#authorizedSearch")) $("#authorizedSearch").addEventListener("input",renderUsers);
    if($("#authorizedStatus")) $("#authorizedStatus").addEventListener("change",renderUsers);
    if($("#authorizedRole")) $("#authorizedRole").addEventListener("change",renderUsers);

    const doLogout=async()=>{
      if(realtimeChannel){await PraxisCloud.unsubscribe(realtimeChannel);realtimeChannel=null;}
      await PraxisCloud.signOut();
      state=null;currentProfile=null;
      showLogin("Sesión cerrada correctamente.","info");
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
      if($("#inventorySearch")) $("#inventorySearch").value=e.target.value;
      page=1;
      if(e.target.value.trim())setView("inventory");
      renderInventory();
    });
  }

  async function init() {
    cloudMode=PraxisCloud.configured();

    // El acceso se conecta primero; así un error del panel nunca bloquea Login/Crear cuenta.
    wireAuthControls();
    wireAppControls();

    if(cloudMode) {
      showLogin();
      try {
        const session=await PraxisCloud.getSession();
        if(session) await enterCloudApp();
      } catch(err) {
        console.error("No se pudo restaurar la sesión",err);
        showLogin((err&&err.message)||"No se pudo conectar con la base de datos.","error");
      }
    } else {
      state=await PraxisDB.get(STATE_KEY);
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
