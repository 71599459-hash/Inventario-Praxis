
const PraxisExcel = (() => {
  const ACTIVE_SHEETS = new Set(["General", "TIC - Almacén", "Cámaras y Redes", "Insum, mat y herra", "Préstamos"]);
  const HISTORIC_SHEETS = new Set(["Compras", "Ventas", "Préstamos"]);
  const SOURCE_PRIORITY = {
    "General": 70, "Préstamos": 62, "TIC - Almacén": 52,
    "Cámaras y Redes": 45, "Insum, mat y herra": 40, "Compras": 20, "Ventas": 10
  };

  const norm = value => String(value ?? "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ").trim().toUpperCase();

  const cleanCode = value => {
    const c = norm(value).replace(/\s+/g, "");
    if (!c || ["#REF!", "-", "N/A", "NA", "NONE", "0", "."].includes(c)) return "";
    if (!/[A-Z]/.test(c) || c.length < 4) return "";
    return c;
  };

  const cleanSite = value => {
    const s = norm(value);
    if (s === "ESTRELAS") return "ESTRELLAS";
    if (["EQUIPO", "-", "0"].includes(s)) return "";
    return s;
  };

  function findHeaderRow(matrix) {
    for (let i = 0; i < Math.min(18, matrix.length); i++) {
      const row = matrix[i].map(norm);
      const score = ["CODIGO TIC","EQUIPO/ MATERIAL","SEDE","AREA / AULA"].filter(x => row.includes(x)).length;
      if (score >= 2) return i;
    }
    return 6;
  }

  function indexMap(headers) {
    const h = headers.map(norm);
    const find = (...names) => {
      for (const name of names) {
        const idx = h.indexOf(norm(name));
        if (idx >= 0) return idx;
      }
      return -1;
    };
    return {
      n:find("N"), bien:find("Bien/Muebles"), sede:find("Sede"), area:find("Área / Aula","Area / Aula"),
      responsable:find("Apellidos y Nombres"), dni:find("DNI"), fechaA:find("Fecha (A)"),
      fechaE:find("Fecha (E)"), fechaI:find("Fecha (I)"), logistica:find("Codigo Logística","Código Logística"),
      padre:find("Código Padre TIC","Codigo Padre TIC"), codigo:find("Código TIC","Codigo TIC"),
      descripcion:find("Descripción","Descripcion"), cantidad:find("Cant."), unidad:find("Unid Med"),
      equipo:find("Equipo/ Material","Equipo / Material"), detalle:find("Detalles equipo"),
      caracteristicas:find("Características técnicas","Caracteristicas tecnicas"), marca:find("Marca"),
      modelo:find("Modelo"), serie:find("Serie o Código","Serie o Codigo"), puertos:find("Puertos"),
      color:find("Color"), estado:find("Estado"), condicion:find("Condición","Condicion"),
      situacion:find("Situación","Situacion"), usadoPor:find("USADO POR"),
      observaciones:find("Observaciones"), revision:find("REVISIÓN","REVISION")
    };
  }

  function parseSheet(ws, sheetName) {
    const matrix = XLSX.utils.sheet_to_json(ws, { header:1, defval:"", raw:false, blankrows:false });
    if (!matrix.length) return [];
    const headerRow = findHeaderRow(matrix);
    const indexes = indexMap(matrix[headerRow] || []);
    const out = [];

    for (let i = headerRow + 1; i < matrix.length; i++) {
      const row = matrix[i] || [];
      const get = key => indexes[key] >= 0 ? String(row[indexes[key]] ?? "").trim() : "";
      const rec = {
        source:sheetName, row:i+1, n:get("n"), bien:get("bien"),
        sede:cleanSite(get("sede")), area:get("area"),
        responsable:get("responsable"), dni:get("dni"),
        fechaA:get("fechaA"), fechaE:get("fechaE"), fechaI:get("fechaI"),
        logistica:get("logistica"), codigoPadre:cleanCode(get("padre")),
        codigo:cleanCode(get("codigo")), descripcion:get("descripcion"),
        cantidad:get("cantidad"), unidad:get("unidad"),
        equipo:norm(get("equipo")) === "0" ? "" : norm(get("equipo")),
        detalle:get("detalle"), caracteristicas:get("caracteristicas"),
        marca:norm(get("marca")), modelo:get("modelo"), serie:get("serie"),
        puertos:get("puertos"), color:get("color"), estado:norm(get("estado")),
        condicion:norm(get("condicion")), situacion:norm(get("situacion")),
        usadoPor:get("usadoPor"), observaciones:get("observaciones") || get("revision")
      };
      const hasData = [rec.codigo,rec.codigoPadre,rec.equipo,rec.sede,rec.area,rec.responsable,rec.marca,rec.modelo,rec.serie].some(Boolean);
      if (hasData) out.push(rec);
    }
    return out;
  }

  function chooseCurrent(records) {
    const score = r => {
      let s = SOURCE_PRIORITY[r.source] || 0;
      if (r.responsable) s += 25;
      if (r.area && !norm(r.area).includes("ALMAC")) s += 12;
      if (norm(r.estado).includes("OPERAT")) s += 3;
      return s;
    };
    return [...records].sort((a,b) => score(b)-score(a))[0];
  }

  function locationType(r) {
    if (r.source === "TIC - Almacén" || norm(r.area).includes("ALMAC")) return "ALMACEN";
    if (r.responsable) return "ASIGNADO";
    return "SEDE";
  }

  function buildState(workbook) {
    const allRecords = [];
    for (const name of workbook.SheetNames) {
      if (!["General","TIC - Almacén","Cámaras y Redes","Insum, mat y herra","Compras","Ventas","Préstamos"].includes(name)) continue;
      allRecords.push(...parseSheet(workbook.Sheets[name], name));
    }

    const historyByCode = new Map();
    for (const r of allRecords) {
      if (!r.codigo) continue;
      if (!historyByCode.has(r.codigo)) historyByCode.set(r.codigo, []);
      historyByCode.get(r.codigo).push({
        type:"EXCEL", source:r.source, row:r.row, sede:r.sede, area:r.area,
        responsable:r.responsable, estado:r.estado,
        fecha:r.fechaI || r.fechaE || r.fechaA || "", observaciones:r.observaciones || ""
      });
    }

    const currentRecords = allRecords.filter(r => ACTIVE_SHEETS.has(r.source));
    const groups = new Map();
    const uncoded = [];
    currentRecords.forEach(r => {
      if (r.codigo) {
        if (!groups.has(r.codigo)) groups.set(r.codigo, []);
        groups.get(r.codigo).push(r);
      } else if (r.equipo || r.codigoPadre) uncoded.push(r);
    });

    const inventory = [];
    for (const [code, records] of groups.entries()) {
      const r = structuredClone(chooseCurrent(records));
      const locationVariants = [...new Set(records.map(x =>
        [x.sede, x.area, x.responsable].map(norm).filter(Boolean).join(" | ")
      ).filter(Boolean))];
      r.id = code;
      r.duplicateSources = records.length;
      r.locationType = locationType(r);
      r.needsReview = locationVariants.length > 1;
      r.conflictLocations = locationVariants;
      r.history = historyByCode.get(code) || [];
      inventory.push(r);
    }
    uncoded.forEach(r0 => {
      const r = structuredClone(r0);
      r.id = `SIN-${r.source.replace(/\W+/g,"-")}-${r.row}`;
      r.duplicateSources = 1;
      r.locationType = locationType(r);
      r.history = [{
        type:"EXCEL", source:r.source, row:r.row, sede:r.sede, area:r.area,
        responsable:r.responsable, estado:r.estado,
        fecha:r.fechaI || r.fechaE || r.fechaA || "", observaciones:r.observaciones || ""
      }];
      inventory.push(r);
    });

    const transactions = allRecords
      .filter(r => HISTORIC_SHEETS.has(r.source))
      .map((r,i) => ({
        id:`TX-${i+1}`, source:r.source, codigo:r.codigo, codigoPadre:r.codigoPadre,
        equipo:r.equipo, sede:r.sede, area:r.area, responsable:r.responsable,
        fecha:r.fechaI || r.fechaE || r.fechaA || "", estado:r.estado,
        observaciones:r.observaciones, cantidad:r.cantidad, marca:r.marca, modelo:r.modelo, serie:r.serie
      }));

    const duplicates = [...groups.values()].filter(v => v.length > 1).length;
    const conflicts = inventory.filter(r => r.needsReview).length;
    return {
      version:2,
      importedAt:new Date().toISOString(),
      workbookSheets:workbook.SheetNames,
      stats:{
        rawRecords:allRecords.length,
        masterRecords:inventory.length,
        duplicateCodes:duplicates,
        conflictCodes:conflicts
      },
      inventory,
      transactions,
      webMovements:[]
    };
  }

  async function fromFile(file) {
    if (!window.XLSX) throw new Error("No se pudo cargar el lector de Excel. Verifica tu conexión a Internet.");
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type:"array", cellDates:true });
    return buildState(workbook);
  }

  return { fromFile, norm, cleanCode };
})();
