const PraxisCargo = (() => {
  const HEAD_TIC = "FABIÁN PUENTE, FRANK JAIME";
  const BASE_SEQUENCE = 16;
  const PAGE_SIZE = 25;

  const esc = value => String(value ?? "").replace(/[&<>"']/g, m => ({
    "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
  }[m]));

  const fmtDate = value => {
    const d = value ? new Date(value + "T12:00:00") : new Date();
    if (Number.isNaN(d.getTime())) return String(value || "");
    return d.toLocaleDateString("es-PE",{day:"2-digit",month:"2-digit",year:"numeric"});
  };

  const personKey = (p, norm) => p.key || `${norm(p.name)}|${String(p.dni||"").trim()}`;

  function nextCode(state) {
    const year = new Date().getFullYear();
    const nums = (state.cargoDocuments || [])
      .map(d => String(d.code || "").match(new RegExp(`^FTEC-${year}-(\\d+)$`,"i")))
      .filter(Boolean)
      .map(m => Number(m[1]))
      .filter(Number.isFinite);
    const seq = Math.max(BASE_SEQUENCE, ...nums) + 1;
    return `FTEC-${year}-${String(seq).padStart(3,"0")}`;
  }

  const itemType = (r, norm) => {
    if (r.bien) return r.bien;
    const e = norm(r.equipo);
    return /(PC|CASE|FUENTE DE PODER|PLACA MADRE|PROCESADOR|DISCO DURO|SSD|RAM)/.test(e) ? "C.INTERNO" : "PERIFÉRICO";
  };

  function snapshot(items) {
    return items
      .slice()
      .sort((a,b)=>(a.codigo||a.id||"").localeCompare(b.codigo||b.id||"","es"))
      .map(r => ({
        id:r.id,
        cantidad:r.cantidad || "1",
        equipo:r.equipo || "",
        bien:r.bien || "",
        marca:r.marca || "",
        modelo:r.modelo || "",
        caracteristicas:r.caracteristicas || r.detalle || r.descripcion || "",
        codigo:r.codigo || "",
        serie:r.serie || "",
        estado:r.estado || "",
        observaciones:r.observaciones || "",
        situacion:r.situacion || ""
      }));
  }

  function softwareHtml() {
    const software = [
      ["Windows","Windows 11",""],["Ms Office Professional Plus","2021",""],
      ["Winrar","",""],["VLC","",""],["Aimp3","",""],
      ["Edge","",""],["Chrome","",""],["Firefox","",""],["Adobe Acrobat","2021",""],["Anydesk","",""]
    ];
    let rows = "";
    for (let row=0; row<5; row++) {
      let cells = "";
      for (let block=0; block<3; block++) {
        const idx = row + block*5;
        const s = software[idx] || ["","",""];
        cells += `<td class="num">${idx+1}</td><td>${esc(s[0])}</td><td>${esc(s[1])}</td><td>${esc(s[2])}</td>`;
      }
      rows += `<tr>${cells}</tr>`;
    }
    return `<div class="section-label">SOFTWARE</div>
      <table class="software-table">
        <thead><tr>
          <th>Ítem</th><th>Nombre</th><th>Versión</th><th>Detalles</th>
          <th>Ítem</th><th>Nombre</th><th>Versión</th><th>Detalles</th>
          <th>Ítem</th><th>Nombre</th><th>Versión</th><th>Detalles</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>`;
  }

  function reviewBox(title, dateValue, techName, observation="") {
    return `<div class="review-box">
      <div class="review-title">${esc(title)}</div>
      <div class="review-date"><b>FECHA</b><span>${dateValue ? esc(fmtDate(dateValue)) : ""}</span></div>
      <div class="review-observation"><b>OBSERVACIONES</b><span>${esc(observation)}</span></div>
      <div class="review-sign">
        <div><div class="sign-space"></div><b>FIRMA DEL USUARIO</b></div>
        <div><div class="sign-space"></div><b>FIRMA DEL TÉCNICO</b></div>
      </div>
      <div class="review-names">
        <div><span>${esc(techName)}</span><b>FIRMA DEL RESPONSABLE DE INF. TEC.</b></div>
        <div><span>${esc(HEAD_TIC)}</span><b>FIRMA DEL JEFE DE TIC</b></div>
      </div>
    </div>`;
  }

  function buildPages(doc, stage, ctx) {
    const techName = String(ctx.currentProfile?.nombre || "CÁRDENAS CURISINCHE, ROBERTO ALEJANDRO").toUpperCase();
    const logo = ctx.logoUrl || "";
    const items = doc.items || [];
    const stageDate = stage === "FINAL" ? (doc.finalDate || ctx.today()) : doc.initialDate;
    const chunks = [];
    const count = Math.max(1, Math.ceil(items.length / PAGE_SIZE));
    for (let i=0; i<count; i++) chunks.push(items.slice(i*PAGE_SIZE,(i+1)*PAGE_SIZE));

    return chunks.map((pageItems,pageIndex) => {
      const globalStart = pageIndex * PAGE_SIZE;
      let rows = "";
      for (let i=0; i<PAGE_SIZE; i++) {
        const r = pageItems[i];
        if (r) {
          const brandModel = [r.marca,r.modelo].filter(Boolean).join(" / ");
          const codeSerie = [r.codigo,r.serie].filter(Boolean).join(" / ");
          rows += `<tr>
            <td class="num">${globalStart+i+1}</td>
            <td class="num">${esc(r.cantidad || "1")}</td>
            <td>${esc(r.equipo)}</td>
            <td>${esc(itemType(r,ctx.norm))}</td>
            <td>${esc(brandModel)}</td>
            <td>${esc(r.caracteristicas)}</td>
            <td>${esc(codeSerie || "SIN CÓDIGO")}</td>
            <td>${esc(r.estado)}</td>
            <td class="nowrap">${esc(fmtDate(stageDate))}</td>
            <td class="signature-row"></td>
            <td class="signature-row"></td>
            <td>${esc(r.observaciones)}</td>
          </tr>`;
        } else {
          rows += `<tr>
            <td class="num">${globalStart+i+1}</td><td></td><td></td><td></td><td></td><td></td>
            <td></td><td></td><td></td><td class="signature-row"></td><td class="signature-row"></td><td></td>
          </tr>`;
        }
      }

      const last = pageIndex === chunks.length - 1;
      const footer = last ? `
        ${softwareHtml()}
        <div class="reviews">
          ${reviewBox("REVISIÓN INICIAL",doc.initialDate,techName,doc.initialObservation||"")}
          ${reviewBox("REVISIÓN FINAL",doc.finalDate,techName,doc.finalObservation||"")}
        </div>` : "";

      return `<section class="print-page">
        <div class="title">FICHA TÉCNICA DE EQUIPO DE CÓMPUTO</div>
        <div class="meta">
          <div><b>Código:</b><span>${esc(doc.code)}</span></div>
          <div><b>Situación:</b><span>${esc(doc.situacion || "INTERNO")}</span></div>
          <div class="wide"><b>Técnico:</b><span>${esc(techName)}</span></div>
          <div><b>Sede:</b><span>${esc(doc.sede || "")}</span></div>
          <div><b>Área:</b><span>${esc(doc.area || "")}</span></div>
          <div class="wide user"><b>Usuario:</b><span>${esc(doc.personName)}${doc.dni ? " · DNI: "+esc(doc.dni) : ""}</span>${logo?`<img src="${logo}">`:""}</div>
        </div>

        <div class="section-label">HARDWARE</div>
        <table class="hardware">
          <thead><tr>
            <th>Item</th><th>Cant.</th><th>Equipo</th><th>Tipo</th><th>Marca / Modelo</th><th>Características</th>
            <th>Código TIC / Serie</th><th>Estado</th><th>Fecha</th><th>Firma (R.C.)</th><th>Firma TIC</th><th>Observación</th>
          </tr></thead>
          <tbody>${rows}</tbody>
        </table>
        ${footer}
        <div class="page-count">Página ${pageIndex+1} de ${chunks.length} · ${stage==="FINAL"?"REVISIÓN FINAL":"ENTREGA INICIAL"}</div>
      </section>`;
    }).join("");
  }

  async function print(person, stage, ctx) {
    if (!ctx.ensureEditable()) return;

    ctx.state.cargoDocuments = ctx.state.cargoDocuments || [];
    const key = personKey(person,ctx.norm);
    let doc = [...ctx.state.cargoDocuments].reverse().find(d => d.personKey===key && !d.finalDate);

    if (stage === "INICIAL") {
      if (!doc) {
        doc = {
          id:crypto.randomUUID(),
          code:nextCode(ctx.state),
          personKey:key,
          personName:person.name,
          dni:person.dni || "",
          sede:[...person.sites].join(", "),
          area:[...person.areas].join(", "),
          situacion:"INTERNO",
          initialDate:ctx.today(),
          finalDate:"",
          initialObservation:"",
          finalObservation:"",
          items:snapshot(person.items),
          createdBy:ctx.currentProfile?.id || "",
          createdByName:ctx.currentProfile?.nombre || ctx.currentProfile?.email || "",
          createdAt:new Date().toISOString()
        };
        ctx.state.cargoDocuments.push(doc);
        await ctx.persist("CARGO_ENTREGA_INICIAL");
      }
    } else {
      if (!doc) {
        ctx.toast("No existe una entrega inicial abierta para este colaborador. Primero imprime la entrega inicial.","error");
        return;
      }
      doc.finalDate = ctx.today();
      doc.finalBy = ctx.currentProfile?.id || "";
      doc.finalByName = ctx.currentProfile?.nombre || ctx.currentProfile?.email || "";
      doc.finalAt = new Date().toISOString();
      await ctx.persist("CARGO_REVISION_FINAL");
    }

    const popup = window.open("","_blank");
    if (!popup) {
      ctx.toast("El navegador bloqueó la impresión. Habilita ventanas emergentes para Inventario Praxis.","error");
      return;
    }

    const pages = buildPages(doc,stage,ctx);
    popup.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(doc.code)} - ${esc(doc.personName)}</title>
    <style>
      @page{size:A4 landscape;margin:5mm}
      *{box-sizing:border-box}body{font-family:Arial,Helvetica,sans-serif;margin:0;color:#111;background:#fff}
      .print-page{width:100%;min-height:198mm;page-break-after:always;position:relative;padding-bottom:3mm}
      .print-page:last-child{page-break-after:auto}
      .title{border:1px solid #222;background:#d9d9d9;text-align:center;font-weight:800;font-size:11pt;padding:4px}
      .meta{display:grid;grid-template-columns:1fr 1fr 2.2fr;border-left:1px solid #222;border-top:1px solid #222;font-size:7pt}
      .meta>div{min-height:18px;display:flex;align-items:center;border-right:1px solid #222;border-bottom:1px solid #222}
      .meta b{width:62px;text-align:right;padding-right:4px}.meta span{flex:1;padding:2px 4px;font-weight:600}
      .meta .user{position:relative;padding-right:86px}.meta .user img{position:absolute;right:6px;top:1px;width:76px;height:31px;object-fit:contain}
      .section-label{background:#d9d9d9;border:1px solid #222;border-top:0;font-size:7.5pt;font-weight:800;padding:2px}
      table{width:100%;border-collapse:collapse;table-layout:fixed}th,td{border:1px solid #222;padding:1.2px 2px;vertical-align:middle;overflow-wrap:anywhere}
      th{background:#d9d9d9;font-size:6.3pt;text-align:cente}.hardware td{font-size:6.1pt;height:15px}
      .hardware th:nth-child(1){width:3%}.hardware th:nth-child(2){width:3%}.hardware th:nth-child(3){width:10%}.hardware th:nth-child(4){width:6%}
      .hardware th:nth-child(5){width:11%}.hardware th:nth-child(6){width:19%}.hardware th:nth-child(7){width:13%}.hardware th:nth-child(8){width:7%}
      .hardware th:nth-child(9){width:7%}.hardware th:nth-child(10){width:7%}.hardware th:nth-child(11){width:6%}.hardware th:nth-child(12){width:8%}
      .num{text-align:center}.nowrap{white-space:nowrap}.signature-row{height:18px}
      .software-table td{font-size:6pt;height:14px}.software-table th:nth-child(4n+1){width:3%}.software-table th:nth-child(4n+2){width:9%}.software-table th:nth-child(4n+3){width:7%}.software-table th:nth-child(4n+4){width:12%}
      .reviews{display:grid;grid-template-columns:1fr 1fr;gap:18mm;margin-top:6mm}
      .review-box{border:1px solid #222;font-size:6.5pt}.review-title{background:#d9d9d9;text-align:center;font-weight:800;border-bottom:1px solid #222;padding:2px}
      .review-date{display:grid;grid-template-columns:42% 58%;border-bottom:1px solid #222;min-height:18px;text-align:center;align-items:center}.review-date b{height:100%;display:grid;place-items:center;border-right:1px solid #222}
      .review-observation{min-height:35px;display:grid;grid-template-rows:14px 1fr;border-bottom:1px solid #2222;text-align:center}.review-observation b{background:#d9d9d9;border-bottom:1px solid #222}.review-observation span{padding:2px}
      .review-sign{display:grid;grid-template-columns:1fr 1fr}.review-sign>div{text-align:center;border-right:1px solid #222}.review-sign>div:last-child{border-right:0}.sign-space{height:40px;border-bottom:1px solid #222}.review-sign b{display:block;padding:2px}
      .review-names{display:grid;grid-template-columns:1fr 1fr;border-top:1px solid #222}.review-names>div{text-align:center;border-right:1px solid #222}.review-names>div:last-child{border-right:0}.review-names span{display:block;min-height:18px;padding:2px;border-bottom:1px solid #222}.review-names b{display:block;padding:2px}
      .page-count{text-align:right;font-size:6pt;color:#555;margin-top:2px}
      @media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
    </style></head><body>${pages}<script>window.onload=()=>setTimeout(()=>window.print(),350)<\/script></body></html>`);
    popup.document.close();
  }

  return {print};
})();
