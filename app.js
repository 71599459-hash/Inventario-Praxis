
(() => {
  const STATE_KEY = "praxis-state-v2";
  const PAGE_SIZE = 30;
  let state = null;
  let currentView = "dashboard";
  let page = 1;
  let selectedId = null;
  let cloudMode = false;
  let currentProfile = null;

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

  function updateProfileUI() {
    const p = currentProfile || PraxisCloud.getProfile();
    if ($("#profileName")) $("#profileName").textContent = (p && (p.nombre || p.email)) || "Usuario";
    if ($("#profileRole")) $("#profileRole").textContent = ((p && p.role) || "LOCAL").replace("_"," ");
    $(".edit-only").forEach(el=>el.classList.toggle("hidden",cloudMode && !canEdit()));
    $(".admin-only").forEach(el=>el.classList.toggle("hidden",!isAdmin()));
    $(".cloud-only").forEach(el=>el.classList.toggle("hidden",!cloudMode));
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
    $("[data-review]").forEach(b=>b.onclick=()=>openAsset(b.dataset.review));
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

    $("[data-pick-conflict]").forEach(b=>b.onclick=async()=>{
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
    if(!ensureEditable()) return;
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
      const r={id:code,codigo:code,codigoPadre:"",equipo:equipment,descripcion:"",marca:norm($("#aBrand").value),modelo:$("#aModel").value.trim(),serie:serial,sede:norm($("#aSite").value),area:$("#aArea").value.trim(),responsable:$("#aPerson").value.trim(),dni:"",estado:$("#aStatus").value,condicion:"",situacion:"INTERNO",source:"Registro web",observaciones:$("#aObs").value.trim(),duplicateSources:1,needsReview:false,conflictLocations:[],conflictRecords:[],locationType:$("#aPerson").value.trim()?"ASIGNADO":norm($("#aArea").value).includes("ALMAC")?"ALMACEN":"SEDE",history:[{type:"WEB",source:"Registro web",fecha:today(),sede:norm($("#aSite").value),area:$("#aArea").value.trim(),responsable:$("#aPerson").value.trim(),estado:$("#aStatus").value,observaciones:"Registro creado desde la web"}]};
      state.inventory.unshift(r);state.webMovements.unshift({id:crypto.randomUUID(),fecha:today(),source:"Alta web",codigo:code,equipo:equipment,to:[r.sede,r.area,r.responsable].filter(Boolean).join(" / "),observaciones:r.observaciones});
      await persist("ALTA_EQUIPO");closeModal();renderAll();toast("Equipo registrado y sincronizado.","success");
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
      const users=await PraxisCloud.listUsers();
      const rows=users.map(function(u){
        return "<tr><td>"+esc(u.nombre||"")+"</td><td>"+esc(u.email||"")+"</td><td><select data-role-user='"+esc(u.id)+"'>"+
          "<option value='ADMIN_TIC' "+(u.role==="ADMIN_TIC"?"selected":"")+">ADMIN_TIC</option>"+
          "<option value='EDITOR' "+(u.role==="EDITOR"?"selected":"")+">EDITOR</option>"+
          "<option value='CONSULTA' "+(u.role==="CONSULTA"?"selected":"")+">CONSULTA</option>"+
          "</select></td><td><span class='badge "+(u.activo?"ok":"bad")+"'>"+(u.activo?"ACTIVO":"INACTIVO")+"</span></td></tr>";
      }).join("");
      $("#usersTable").innerHTML="<div class='table-wrap'><table><thead><tr><th>Nombre</th><th>Correo</th><th>Rol</th><th>Estado</th></tr></thead><tbody>"+(rows||"<tr><td colspan='4'>Sin usuarios.</td></tr>")+"</tbody></table></div>";
      $("[data-role-user]").forEach(function(sel){
        sel.onchange=async function(){
          const role=sel.value;
          if(!confirm("¿Cambiar el rol a "+role+"?")){ await renderUsers(); return; }
          try {
            await PraxisCloud.setUserRole(sel.dataset.roleUser,role);
            toast("Rol actualizado.","success");
            await renderUsers();
            await renderAudit();
          } catch(err) {
            console.error(err);
            toast((err&&err.message)||"No se pudo actualizar el rol.","error");
            await renderUsers();
          }
        };
      });
    } catch(err) {
      console.error(err);
      $("#usersTable").innerHTML="<div class='alert error'>"+esc((err&&err.message)||"No se pudieron cargar los usuarios.")+"</div>";
    }
  }

  async function renderAudit() {
    if(!isAdmin() || !$("#auditTable")) return;
    try {
      const items=await PraxisCloud.audit(100);
      const rows=items.map(function(a){
        return "<tr><td>"+esc(new Date(a.created_at).toLocaleString("es-PE"))+"</td><td>"+esc(a.action)+"</td><td>"+esc(a.user_id||"")+"</td><td>"+esc(a.version_before==null?"":a.version_before)+"</td><td>"+esc(a.version_after==null?"":a.version_after)+"</td></tr>";
      }).join("");
      $("#auditTable").innerHTML="<div class='table-wrap'><table><thead><tr><th>Fecha</th><th>Acción</th><th>Usuario</th><th>Versión anterior</th><th>Nueva versión</th></tr></thead><tbody>"+(rows||"<tr><td colspan='5'>Sin eventos.</td></tr>")+"</tbody></table></div>";
    } catch(err) {
      console.error(err);
      $("#auditTable").innerHTML="<div class='alert error'>"+esc((err&&err.message)||"No se pudo cargar la auditoría.")+"</div>";
    }
  }

  async function enterCloudApp() {
    currentProfile=await PraxisCloud.loadProfile();
    if(!currentProfile || !currentProfile.activo) throw new Error("Tu usuario está inactivo.");
    const remote=await PraxisCloud.loadState();
    state=remote.state;
    if(state) await PraxisDB.set(STATE_KEY,state);
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
    cloudMode=PraxisCloud.configured();

    $(".nav-item").forEach(b=>b.onclick=()=>setView(b.dataset.view));
    $("[data-go]").forEach(b=>b.onclick=()=>setView(b.dataset.go));
    $("#mobileMenu").onclick=()=>$("#sidebar").classList.toggle("open");
    $("#modalClose").onclick=closeModal;
    $("#modal").addEventListener("click",e=>{if(e.target.id==="modal")closeModal()});
    $("#dismissPrivacy").onclick=()=>$("#privacyBanner").remove();
    $("#importBtnTop").onclick=()=>$("#excelInput").click();
    $("#importBtnEmpty").onclick=()=>$("#excelInput").click();
    $("#excelInput").onchange=e=>{const f=e.target.files[0];if(f)importExcel(f);e.target.value=""};
    $("#demoInfoBtn").onclick=demoInfo;
    $("#addAssetBtn").onclick=openAddAsset;
    $("#exportCsvBtn").onclick=exportInventory;
    $("#backupBtn").onclick=downloadBackup;
    $("#exportMovementsBtn").onclick=exportMovements;
    $("#restoreBtn").onclick=()=>$("#restoreInput").click();
    $("#restoreInput").onchange=e=>{const f=e.target.files[0];if(f)restoreBackup(f);e.target.value=""};
    $("#clearLocalBtn").onclick=clearLocalData;
    if($("#refreshUsersBtn")) $("#refreshUsersBtn").onclick=renderUsers;
    if($("#refreshAuditBtn")) $("#refreshAuditBtn").onclick=renderAudit;
    if($("#logoutBtn")) $("#logoutBtn").onclick=async()=>{
      await PraxisCloud.signOut();
      state=null;currentProfile=null;
      showAuth("Sesión cerrada.","info");
    };

    ["#inventorySearch","#siteFilter","#categoryFilter","#statusFilter","#locationFilter"].forEach(sel=>$(sel).addEventListener("input",()=>{page=1;renderInventory()}));
    $("#peopleSearch").addEventListener("input",renderPeople);
    $("#reviewSearch").addEventListener("input",renderReview);
    $("#globalSearch").addEventListener("input",e=>{
      if(!state?.inventory?.length)return;
      $("#inventorySearch").value=e.target.value;
      page=1;
      if(e.target.value.trim())setView("inventory");
      renderInventory();
    });

    if($("#loginForm")) $("#loginForm").onsubmit=async e=>{
      e.preventDefault();
      const email=$("#loginEmail").value.trim();
      const password=$("#loginPassword").value;
      const box=$("#authMessage");
      box.textContent="Ingresando…";
      box.className="auth-message show info";
      try {
        await PraxisCloud.signIn(email,password);
        await enterCloudApp();
      } catch(err) {
        console.error(err);
        showAuth((err&&err.message)||"No se pudo iniciar sesión.","error");
      }
    };

    if(cloudMode) {
      try {
        const session=await PraxisCloud.getSession();
        if(session) await enterCloudApp();
        else showAuth();
      } catch(err) {
        console.error(err);
        showAuth("No se pudo conectar con la base de datos. Revisa la configuración de Supabase.","error");
      }
    } else {
      state=await PraxisDB.get(STATE_KEY);
      $("#authGate").classList.add("hidden");
      $("#appShell").classList.remove("hidden");
      updateProfileUI();
      renderAll();
    }
  }

  document.addEventListener("DOMContentLoaded",init);
})();
