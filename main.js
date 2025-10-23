/* main.js - Integrado con UI de Ejecutivo (engranaje + modal) y sin alterar la lógica existente
   Comentarios en tercera persona impersonal, pensados para un lector con nivel básico-intermedio.
*/

/* ------------ Estado global ------------
   - Se mantienen arrays para el catálogo y la estructura de UI.
   - state almacena el estado dinámico de secciones y subsecciones.
*/
let catalog = [];
let structure = [];
const state = { sections: {} };

/* ------------ Inicialización ------------
   - Se registra el evento DOMContentLoaded para cargar datos, inicializar contrato y la UI del ejecutivo.
*/
document.addEventListener('DOMContentLoaded', () => {
  // Se intenta cargar data.xlsx desde la raíz; si falla, se muestra mensaje de error.
  loadFromServer('data.xlsx').catch(err => {
    console.error('Carga inicial falló:', err);
    showMessage('No se encontró data.xlsx en la raíz o hubo un error al leerlo. Asegúrate de servir la app vía HTTP y de que data.xlsx exista en la raíz.', true, 0);
  });
  // Se configura la lógica de contrato (formularios, generación de docx/pdf).
  inicializarContrato();
  // Se inicializa la UI y persistencia del Ejecutivo (modal, botones).
  initEjecutivoUI();
});

/* ------------ Mensajes visibles ------------
   - showMessage: muestra mensajes en la UI (color y duración configurables).
*/
function showMessage(msg, isError = true, timeout = 8000) {
  const el = document.getElementById('messages');
  if (!el) return;
  el.textContent = msg;
  el.style.color = isError ? '#b71c1c' : '#1b5e20';
  if (timeout) setTimeout(() => { if (el.textContent === msg) el.textContent = ''; }, timeout);
}

/* ------------ Cargar data.xlsx desde raíz ------------
   - loadFromServer: descarga el archivo y lo pasa a parseWorkbook.
*/
async function loadFromServer(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const ab = await res.arrayBuffer();
  parseWorkbook(ab);
}

/* ------------ Parseo y normalización ------------
   - parseWorkbook: lee el workbook XLSX, mapea columnas y normaliza campos importantes.
   - Se soportan nombres alternativos de columnas y valores por defecto.
*/
function parseWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });

  // Se busca la hoja "catalog" (insensible a mayúsculas) o se usa la primera hoja.
  const catalogSheetName = workbook.SheetNames.find(n => n.toLowerCase() === 'catalog') || workbook.SheetNames[0];
  const rawCatalogRows = XLSX.utils.sheet_to_json(workbook.Sheets[catalogSheetName], { defval: '' });

  // Se mapea cada fila a un objeto con campos estandarizados.
  catalog = rawCatalogRows.map(row => {
    const mapped = {
      Código: row['Código'] || row['Codigo'] || row['Code'] || '',
      Plan: row['Plan'] || row['Name'] || '',
      Valor: row['Valor'] || row['Value'] || row['Price'] || '',
      Promo1: row['Promo1'] || row['Promo_1'] || '',
      Meses1: row['Meses1'] || row['Meses_1'] || '',
      Promo2: row['Promo2'] || row['Promo_2'] || '',
      Meses2: row['Meses2'] || row['Meses_2'] || '',
      Detalles: row['Detalles'] || row['Details'] || '',
      Section: row['Section'] || row['Sección'] || row['Seccion'] || '',
      Subsection: row['Subsection'] || row['Subsección'] || row['Subseccion'] || '',
      ExtraFor: row['ExtraFor'] || row['Extra_for'] || ''
    };
    // Se normalizan números para evitar problemas de formato.
    mapped.Valor = normalizeNumber(mapped.Valor);
    mapped.Promo1 = normalizeNumber(mapped.Promo1);
    mapped.Promo2 = normalizeNumber(mapped.Promo2);
    // Meses se permiten como texto en algunos casos (ej. "12" o "12 meses")
    mapped.Meses1 = normalizeNumber(mapped.Meses1, true, true);
    mapped.Meses2 = normalizeNumber(mapped.Meses2, true, true);
    return mapped;
  });

  // Se verifica si faltan columnas obligatorias y se advierte si aplica.
  const firstRaw = rawCatalogRows[0] || {};
  const headerKeys = Object.keys(firstRaw).map(k => String(k).toLowerCase());
  const requiredCols = ['código','plan','valor','promo1','meses1','promo2','meses2','detalles'];
  const missing = requiredCols.filter(c => !headerKeys.includes(c));
  if (missing.length) {
    console.warn('Encabezados faltantes detectados en catalog:', missing);
    showMessage(`Advertencia: faltan encabezados obligatorios en catalog: ${missing.join(', ')}. El parser intentará mapear columnas alternativas.`, true, 12000);
  }

  // Se intenta leer la hoja "structure" para configurar la UI; si no existe, se usa estructura por defecto.
  const structureSheetName = workbook.SheetNames.find(n => n.toLowerCase() === 'structure');
  if (structureSheetName) {
    const rawStruct = XLSX.utils.sheet_to_json(workbook.Sheets[structureSheetName], { defval: '' });
    structure = rawStruct.map(r => ({
      Section: r['Section'] || r['Sección'] || r['Seccion'] || '',
      Subsection: r['Subsection'] || r['Subsección'] || r['Subseccion'] || '',
      ComponentType: r['ComponentType'] || r['Tipo'] || '',
      Prefixes: r['Prefixes'] || '',
      ToggleOptions: r['ToggleOptions'] || '',
      MultiPrefixes: r['MultiPrefixes'] || '',
      MaxAdditional: Number(r['MaxAdditional'] || r['MaxAdicional'] || 4),
      ExtraMapping: r['ExtraMapping'] || ''
    }));
  } else {
    // Estructura por defecto en caso de ausencia de la hoja "structure".
    structure = [
      { Section: 'Hogar', Subsection: 'trio', ComponentType: 'trio', Prefixes: 'T', MaxAdditional: 0 },
      { Section: 'Hogar', Subsection: 'duo', ComponentType: 'duo', ToggleOptions: 'fibra_tv:DT,fibra_fijo:DF,tv_fijo:DTF', MaxAdditional: 0 },
      { Section: 'Hogar', Subsection: 'uno', ComponentType: 'uno', ToggleOptions: 'fibra:F,tv:TV,fijo:FI', MaxAdditional: 0 },
      { Section: 'Movil', Subsection: 'nuevo', ComponentType: 'movil_group', MultiPrefixes: 'multi:NM,datos:ND,voz:NV', MaxAdditional: 4, ExtraMapping: 'NM02:NM02S;NM03:NM03S' },
      { Section: 'Movil', Subsection: 'cartera', ComponentType: 'movil_group', MultiPrefixes: 'multi:CM,datos:CD,voz:CV', MaxAdditional: 4, ExtraMapping: 'CM02:CM02S;CM03:CM03S' }
    ];
    showMessage('Hoja "structure" no encontrada: uso estructura por defecto.', false, 6000);
  }

  console.info('Catalog rows (muestra 10):', catalog.slice(0,10));
  // Se construye la UI dinámica a partir de la estructura procesada.
  buildUIFromStructure();
}

/* ------------ Normalización de números ------------
   - normalizeNumber: intenta extraer un número desde diferentes formatos (ej. "1.234,56" o "1234.56").
   - Parámetros:
     - val: valor a normalizar.
     - integer: si true, se devuelve entero.
     - allowText: si true y no es posible parsear número, devuelve la cadena original.
*/
function normalizeNumber(val, integer = false, allowText = false) {
  if (val === null || val === undefined || val === '') return allowText ? '' : '';
  if (typeof val === 'number') return integer ? Math.floor(val) : val;

  let s = String(val).trim();
  if (!s) return allowText ? '' : '';

  const low = s.toLowerCase();
  if (['no aplica','noaplica','n/a','na','-'].includes(low)) return '';

  const match = s.match(/-?\d[\d\.\,]*/);
  if (match) {
    let numStr = match[0];
    // Se manejan distintos separadores de miles/decimales.
    if (numStr.indexOf('.') !== -1 && numStr.indexOf(',') !== -1) {
      numStr = numStr.replace(/\./g, '').replace(',', '.');
    } else if (numStr.indexOf(',') !== -1 && numStr.indexOf('.') === -1) {
      numStr = numStr.replace(',', '.');
    } else {
      const dots = (numStr.match(/\./g) || []).length;
      if (dots > 1) numStr = numStr.replace(/\./g, '');
    }

    const num = Number(numStr);
    if (!Number.isNaN(num)) return integer ? Math.floor(num) : num;
  }

  if (allowText) return s;
  return '';
}

/* ------------ matchesPrefix ------------
   - Comprueba si un código comienza con un prefijo válido.
   - Evita falsos positivos cuando el siguiente carácter es letra (p. ej. prefijo "NM" y código "NMX" podrían tratarse distinto).
*/
function matchesPrefix(code = '', prefix = '') {
  if (!code || !prefix) return false;
  code = String(code);
  prefix = String(prefix);
  if (code === prefix) return true;
  if (!code.startsWith(prefix)) return false;
  const next = code.charAt(prefix.length);
  if (!next) return true;
  return !(/[A-Za-z]/.test(next));
}

/* ------------ Helpers para limpiar selecciones ------------
   - Funciones que reinician selects y campos de portabilidad.
*/
function resetSectionSelections(sectionName) {
  const sec = state.sections[sectionName];
  if (!sec || !sec.subsections) return;
  Object.keys(sec.subsections).forEach(subName => {
    const st = sec.subsections[subName];
    if (!st || !st.elementos) return;
    if (st.elementos.mainSelect) {
      try { st.elementos.mainSelect.value = ''; } catch (e) {}
    }
    if (Array.isArray(st.elementos.lines)) {
      st.elementos.lines.forEach(line => {
        try {
          if (line.select) line.select.value = '';
          if (line.portaCheckbox) {
            line.portaCheckbox.checked = false;
            if (line.portaFields) line.portaFields.style.display = 'none';
          }
          if (line.portaNumeroInput) line.portaNumeroInput.value = '';
          if (line.portaDonanteInput) line.portaDonanteInput.value = '';
        } catch (e) {}
      });
    }
  });
}

function clearSelectionsExceptSection(keepSection) {
  Object.keys(state.sections).forEach(secName => {
    if (secName !== keepSection) resetSectionSelections(secName);
  });
}

function resetSubsectionsExcept(sectionName, keepSub) {
  const sec = state.sections[sectionName];
  if (!sec || !sec.subsections) return;
  Object.keys(sec.subsections).forEach(subName => {
    if (subName === keepSub) return;
    const st = sec.subsections[subName];
    if (!st || !st.elementos) return;
    if (st.elementos.mainSelect) {
      try { st.elementos.mainSelect.value = ''; } catch (e) {}
    }
    if (Array.isArray(st.elementos.lines)) {
      st.elementos.lines.forEach(line => {
        try {
          if (line.select) line.select.value = '';
          if (line.portaCheckbox) {
            line.portaCheckbox.checked = false;
            if (line.portaFields) line.portaFields.style.display = 'none';
          }
          if (line.portaNumeroInput) line.portaNumeroInput.value = '';
          if (line.portaDonanteInput) line.portaDonanteInput.value = '';
        } catch (e) {}
      });
    }
  });
}

/* ------------ Construcción dinámica de UI ------------
   - buildUIFromStructure: crea pestañas principales y subsecciones a partir de 'structure'.
   - renderSection / renderSubsectionContent: renderizan la UI según el tipo de componente.
*/
function buildUIFromStructure() {
  const bySection = {};
  structure.forEach(s => { if (!bySection[s.Section]) bySection[s.Section] = []; bySection[s.Section].push(s); });

  const root = document.getElementById('tarifario-root');
  root.innerHTML = '';

  // Se crean pestañas por sección
  const tabs = document.createElement('div'); tabs.className = 'tabs';
  const sectionNames = Object.keys(bySection);
  sectionNames.forEach((sec, i) => {
    const btn = document.createElement('button');
    btn.className = 'tab-btn' + (i === 0 ? ' active' : '');
    btn.textContent = sec;
    btn.dataset.section = sec;
    tabs.appendChild(btn);
    btn.addEventListener('click', () => {
      clearSelectionsExceptSection(sec);
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      renderSection(sec, bySection[sec]);
    });
    // Se inicializa el estado de la sección
    state.sections[sec] = { subsections: {}, activeSub: null };
  });
  root.appendChild(tabs);

  const panel = document.createElement('div'); panel.className = 'section-panel';
  root.appendChild(panel);

  // Se muestra la primera sección por defecto
  if (sectionNames.length) renderSection(sectionNames[0], bySection[sectionNames[0]]);
}

function renderSection(sectionName, subsections) {
  const panel = document.querySelector('#tarifario-root .section-panel');
  panel.innerHTML = '';

  // Se crean subtabs (subsecciones) y área de contenido
  const subtabs = document.createElement('div'); subtabs.className = 'subtabs';
  panel.appendChild(subtabs);

  const contentArea = document.createElement('div');
  panel.appendChild(contentArea);

  subsections.forEach((sub, idx) => {
    const sbtn = document.createElement('button');
    sbtn.className = 'subtab-btn' + (idx === 0 ? ' active' : '');
    sbtn.textContent = sub.Subsection;
    sbtn.dataset.sub = sub.Subsection;
    subtabs.appendChild(sbtn);

    // Se prepara el objeto de estado para la subsección
    state.sections[sectionName].subsections[sub.Subsection] = { config: sub, elementos: {}, main: {} };

    sbtn.addEventListener('click', () => {
      resetSubsectionsExcept(sectionName, sub.Subsection);
      subtabs.querySelectorAll('.subtab-btn').forEach(b => b.classList.remove('active'));
      sbtn.classList.add('active');
      renderSubsectionContent(contentArea, sectionName, sub);
      state.sections[sectionName].activeSub = sub.Subsection;
    });

    // Se renderiza la primera subsección por defecto
    if (idx === 0) {
      renderSubsectionContent(contentArea, sectionName, sub);
      state.sections[sectionName].activeSub = sub.Subsection;
    }
  });
}

function renderSubsectionContent(container, sectionName, subCfg) {
  // Se limpia el contenedor y se crean columnas izquierda/derecha
  container.innerHTML = '';
  const row = document.createElement('div'); row.className = 'row';
  const colLeft = document.createElement('div'); colLeft.className = 'col';
  const colRight = document.createElement('div'); colRight.className = 'col';
  row.appendChild(colLeft); row.appendChild(colRight);
  container.appendChild(row);

  const stateSub = state.sections[sectionName].subsections[subCfg.Subsection];

  // Helper para obtener opciones según prefijos
  const optionsFromPrefixes = (prefixStr) => {
    if (!prefixStr) return [];
    const prefs = prefixStr.split(',').map(s => s.trim()).filter(Boolean);
    return catalog.filter(r => prefs.some(p => matchesPrefix(String(r.Código || ''), p)));
  };

  // Se renderiza según el tipo de componente definido en structure
  if (subCfg.ComponentType === 'trio') {
    // Select simple para plan tipo 'trio'
    colLeft.appendChild(createLabel(`Plan ${subCfg.Subsection}`));
    const sel = document.createElement('select');
    sel.innerHTML = `<option value="">-- Selecciona --</option>`;
    const opts = optionsFromPrefixes(subCfg.Prefixes || 'T');
    opts.forEach(o => sel.insertAdjacentHTML('beforeend', `<option value="${o.Código}">${o.Plan}</option>`));
    sel.addEventListener('change', () => {
      if (!sel.value) updateDetalleYPreciosDefault(colRight); else updateDetalleYPrecios(colRight, findByCode(sel.value));
      stateSub.main = { codigo: sel.value };
    });
    colLeft.appendChild(sel);
    stateSub.elementos.mainSelect = sel;
    colRight.appendChild(createOfferBox('Selecciona un plan para ver detalles.'));
  } else if (subCfg.ComponentType === 'duo' || subCfg.ComponentType === 'uno') {
    // Toggle + select: permite cambiar subconjunto de prefijos mediante botones
    colLeft.appendChild(createLabel(`${subCfg.Subsection} - Opciones`));
    const segDiv = document.createElement('div'); segDiv.className = 'segmented-toggle';
    colLeft.appendChild(segDiv);

    const sel = document.createElement('select'); sel.innerHTML = `<option value="">-- Selecciona --</option>`;
    colLeft.appendChild(sel);
    stateSub.elementos.mainSelect = sel;

    const pairs = parseToggleOptions(subCfg.ToggleOptions || '');
    if (pairs.length === 0) {
      const note = document.createElement('div'); note.textContent = 'No hay opciones de toggle definidas en structure';
      colLeft.appendChild(note);
    } else {
      pairs.forEach((p, i) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        const labelText = String(p.key).replace(/[_\-]/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
        btn.textContent = labelText;
        if (i === 0) btn.classList.add('active');
        btn.addEventListener('click', () => {
          segDiv.querySelectorAll('button').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          populateSelectFromPrefixes(sel, p.prefixes);
          sel.value = '';
          updateDetalleYPreciosDefault(colRight);
        });
        segDiv.appendChild(btn);
      });
      // Se carga la primera opción por defecto
      populateSelectFromPrefixes(sel, pairs[0].prefixes);
    }

    sel.addEventListener('change', () => {
      if (!sel.value) updateDetalleYPreciosDefault(colRight); else updateDetalleYPrecios(colRight, findByCode(sel.value));
      stateSub.main = { codigo: sel.value };
    });
    colRight.appendChild(createOfferBox('Selecciona un plan para ver detalles.'));
  } else if (subCfg.ComponentType === 'movil_group') {
    // Grupo móvil: se permite añadir líneas (principal + adicionales)
    const wrapper = document.createElement('div'); wrapper.className = 'movil-lines';
    colLeft.appendChild(wrapper);

    const principal = createMovilLine(sectionName, subCfg, 0, false);
    wrapper.appendChild(principal.lineElement);
    stateSub.elementos.lines = [principal];
    stateSub.main = { lines: [{ idx: 0, codigo: '' }] };

    const addBtn = document.createElement('button'); addBtn.type = 'button'; addBtn.className = 'btn btn-add';
    addBtn.textContent = '+ Añadir línea';
    addBtn.addEventListener('click', () => {
      const max = Math.max(0, subCfg.MaxAdditional || 4);
      const currentAdditional = stateSub.elementos.lines.length - 1;
      if (currentAdditional >= max) { alert(`Máximo ${max} líneas adicionales`); return; }
      const newIdx = stateSub.elementos.lines.length;
      const newLine = createMovilLine(sectionName, subCfg, newIdx, true);
      stateSub.elementos.lines.push(newLine);
      wrapper.appendChild(newLine.lineElement);
      actualizarMovilSection(sectionName, subCfg.Subsection);
    });
    colLeft.appendChild(addBtn);

    // Contenedores de detalles y precios que se actualizan dinámicamente
    const detallesBox = document.createElement('div'); detallesBox.className = 'offer-details';
    detallesBox.innerHTML = 'Selecciona las líneas para ver detalles.';
    const preciosBox = document.createElement('div'); preciosBox.className = 'precios';
    const factBox = document.createElement('div'); factBox.className = 'facturacion';
    colRight.appendChild(detallesBox); colRight.appendChild(preciosBox); colRight.appendChild(factBox);

    stateSub.elementos.detallesBox = detallesBox;
    stateSub.elementos.preciosBox = preciosBox;
    stateSub.elementos.facturacionBox = factBox;
  } else {
    // Componente genérico: select simple
    colLeft.appendChild(createLabel(`Plan ${subCfg.Subsection}`));
    const sel = document.createElement('select');
    sel.innerHTML = `<option value="">-- Selecciona --</option>`;
    const opts = optionsFromPrefixes(subCfg.Prefixes || '');
    opts.forEach(o => sel.insertAdjacentHTML('beforeend', `<option value="${o.Código}">${o.Plan}</option>`));
    sel.addEventListener('change', () => {
      if (!sel.value) updateDetalleYPreciosDefault(colRight); else updateDetalleYPrecios(colRight, findByCode(sel.value));
      stateSub.main = { codigo: sel.value };
    });
    colLeft.appendChild(sel);
    colRight.appendChild(createOfferBox('Selecciona un plan para ver detalles.'));
    stateSub.elementos.mainSelect = sel;
  }
}

/* ------------ UI helpers y movil line creation ------------ */
function createLabel(text) { const l = document.createElement('label'); l.textContent = text; l.style.display = 'block'; l.style.marginTop = '6px'; return l; }

// createOfferBox: caja visual para mostrar detalles iniciales
function createOfferBox(initialText) {
  const box = document.createElement('div');
  box.className = 'offer-details';
  const txt = document.createElement('div');
  txt.className = 'offer-details-text';
  txt.textContent = initialText;
  box.appendChild(txt);
  return box;
}

// findByCode: busca un plan en el catálogo por su código
function findByCode(code) { return catalog.find(r => r.Código === code) || null; }

/* updateDetalleYPrecios / updateDetalleYPreciosDefault
   - Se actualizan los bloques de detalles y precios en la columna derecha según el plan seleccionado.
*/
function updateDetalleYPrecios(containerColRight, plan) {
  containerColRight.innerHTML = '';
  const box = document.createElement('div'); box.className = 'offer-details';
  const title = document.createElement('div'); title.innerHTML = plan ? `<strong>${escapeHtml(plan.Plan)}</strong>` : '<strong>Plan</strong>';
  const details = document.createElement('div'); details.className = 'offer-details-text';
  details.textContent = plan ? (plan.Detalles || '') : 'Selecciona un plan para ver detalles.';
  box.appendChild(title);
  box.appendChild(details);

  if (plan) {
    const precios = document.createElement('div'); precios.className = 'precios';
    const lines = [];
    if (plan.Promo1 !== '' && plan.Promo1 !== null && plan.Promo1 !== undefined) {
      const months1Text = (typeof plan.Meses1 === 'number') ? `${plan.Meses1} meses` : (plan.Meses1 || '-');
      lines.push(`Promo 1: $${plan.Promo1} (${months1Text})`);
    }
    if (plan.Promo2 !== '' && plan.Promo2 !== null && plan.Promo2 !== undefined) {
      const months2Text = (typeof plan.Meses2 === 'number') ? `${plan.Meses2} meses` : (plan.Meses2 || '-');
      lines.push(`Promo 2: $${plan.Promo2} (${months2Text})`);
    }
    if (plan.Valor !== '' && plan.Valor !== null && plan.Valor !== undefined) {
      lines.push(`Sin descuento: $${plan.Valor}`);
    }
    precios.innerHTML = lines.map(l => `<div>${escapeHtml(l)}</div>`).join('');
    box.appendChild(precios);
  } else {
    const defaultBox = createOfferBox('Selecciona un plan para ver detalles.');
    containerColRight.appendChild(defaultBox);
    return;
  }
  containerColRight.appendChild(box);
}
function updateDetalleYPreciosDefault(containerColRight) {
  containerColRight.innerHTML = '';
  const box = createOfferBox('Selecciona un plan para ver detalles.');
  containerColRight.appendChild(box);
}

// parseToggleOptions: parsea opciones tipo "key:prefix1|prefix2,..." en estructura legible
function parseToggleOptions(str) {
  if (!str) return [];
  return str.split(',').map(s => {
    const [k, pref] = s.split(':').map(x => x && x.trim());
    return { key: k || '', prefixes: (pref || '').split('|').map(p => p.trim()).filter(Boolean) };
  }).filter(x => x.key);
}

// populateSelectFromPrefixes: rellena un select con planes que coincidan con una lista de prefijos
function populateSelectFromPrefixes(selectEl, prefixes) {
  selectEl.innerHTML = `<option value="">-- Selecciona --</option>`;
  if (!prefixes || prefixes.length === 0) return;
  const opts = catalog.filter(r => prefixes.some(p => matchesPrefix(String(r.Código || ''), String(p || '').trim())));
  opts.forEach(o => selectEl.insertAdjacentHTML('beforeend', `<option value="${o.Código}">${escapeHtml(o.Plan)}</option>`));
}

/* ------------ Movil: creación de línea (tarjeta) ------------
   - createMovilLine: crea la UI de una línea (principal o adicional) con select, toggle pequeño y campos de portabilidad.
*/
function createMovilLine(sectionName, cfg, idx, esSecundaria) {
  const card = document.createElement('div');
  card.className = 'movil-line-card';

  // Header con título y toggle pequeño
  const header = document.createElement('div');
  header.className = 'movil-line-header';

  const title = document.createElement('div');
  title.className = 'movil-line-title';
  title.textContent = esSecundaria ? `Adicional ${idx}` : 'Línea Principal';

  const headerRight = document.createElement('div');
  headerRight.className = 'movil-line-header-right';

  const segSmall = document.createElement('div'); segSmall.className = 'segmented-toggle small';
  const estados = ['multi','datos','voz'];
  estados.forEach((e,i) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.textContent = e.charAt(0).toUpperCase() + e.slice(1);
    if (i === 0) btn.classList.add('active');
    btn.addEventListener('click', () => {
      // Al cambiar estado, se recargan opciones en el select según la configuración.
      segSmall.querySelectorAll('button').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      cargarOpcionesMovilDesdeCfg(sectionName, cfg, select, e, idx, esSecundaria);
      if (select.value) select.dispatchEvent(new Event('change'));
      else {
        const stateSub = findStateSub(sectionName, cfg.Subsection);
        if (stateSub && stateSub.elementos && stateSub.elementos.detallesBox) {
          stateSub.elementos.detallesBox.innerHTML = `<div><strong>Línea ${idx === 0 ? 'Principal' : 'Adicional ' + idx}</strong><br>Selecciona un plan para ver detalles.</div>`;
        }
        actualizarMovilSection(sectionName, cfg.Subsection);
      }
    });
    segSmall.appendChild(btn);
  });

  headerRight.appendChild(segSmall);

  header.appendChild(title);
  header.appendChild(headerRight);
  card.appendChild(header);

  // Content: select + porta
  const content = document.createElement('div'); content.className = 'movil-line-content';

  const select = document.createElement('select');
  select.id = `${sectionName}-${cfg.Subsection}-select-${idx}`;
  select.innerHTML = `<option value="">-- Selecciona --</option>`;

  // Controles de portabilidad (checkbox + campos opcionales)
  const portaContainer = document.createElement('div'); portaContainer.className = 'porta-container';
  const portaLabel = document.createElement('label'); portaLabel.className = 'porta-label';
  const portaCheckbox = document.createElement('input');
  portaCheckbox.type = 'checkbox';
  portaCheckbox.className = 'porta-checkbox';
  portaLabel.appendChild(portaCheckbox);
  const portaText = document.createElement('span'); portaText.textContent = ' Porta';
  portaLabel.appendChild(portaText);
  portaContainer.appendChild(portaLabel);

  const portaFields = document.createElement('div'); portaFields.classList.add('porta-fields');
  portaFields.style.display = 'none';
  const inputNumero = document.createElement('input'); inputNumero.type = 'text'; inputNumero.placeholder = 'Número a portar';
  inputNumero.className = 'porta-numero';
  const inputDonante = document.createElement('input'); inputDonante.type = 'text'; inputDonante.placeholder = 'Compañía donante';
  inputDonante.className = 'porta-donante';
  portaFields.appendChild(inputNumero);
  portaFields.appendChild(inputDonante);

  portaContainer.appendChild(portaFields);

  // Mostrar/ocultar campos de portabilidad al marcar el checkbox
  portaCheckbox.addEventListener('change', () => {
    portaFields.style.display = portaCheckbox.checked ? 'block' : 'none';
    actualizarMovilSection(sectionName, cfg.Subsection);
  });

  // Al cambiar el select, se recalcula la sección móvil
  select.addEventListener('change', () => {
    if (!select.value) {
      const stateSub = findStateSub(sectionName, cfg.Subsection);
      if (stateSub && stateSub.elementos && stateSub.elementos.detallesBox) {
        stateSub.elementos.detallesBox.innerHTML = `<div><strong>Línea ${idx === 0 ? 'Principal' : 'Adicional ' + idx}</strong><br>Selecciona un plan para ver detalles.</div>`;
      }
      actualizarMovilSection(sectionName, cfg.Subsection);
    } else {
      actualizarMovilSection(sectionName, cfg.Subsection);
    }
  });

  content.appendChild(select);
  content.appendChild(portaContainer);
  card.appendChild(content);

  // Cargar opciones iniciales (estado 'multi' por defecto)
  cargarOpcionesMovilDesdeCfg(sectionName, cfg, select, 'multi', idx, esSecundaria);

  let removeBtn = null;
  if (esSecundaria) {
    // Botón para eliminar líneas adicionales
    removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'btn btn-remove';
    removeBtn.textContent = '🗑';
    removeBtn.title = 'Eliminar línea';
    removeBtn.addEventListener('click', () => {
      const stateSub = findStateSub(sectionName, cfg.Subsection);
      if (!stateSub) return;
      const arr = stateSub.elementos.lines;
      const pos = arr.findIndex(l => l.lineElement === card);
      if (pos >= 0) arr.splice(pos, 1);
      card.remove();
      // Re-indexar títulos de líneas adicionales
      arr.slice(1).forEach((l, i) => {
        const newIdx = i+1;
        l.lineElement.dataset.idx = newIdx;
        const lbl = l.lineElement.querySelector('.movil-line-title');
        if (lbl) lbl.textContent = `Adicional ${newIdx}`;
      });
      actualizarMovilSection(sectionName, cfg.Subsection);
    });
    card.appendChild(removeBtn);
  }

  // Se devuelve un objeto con los elementos relevantes para mantener el estado
  return {
    idx,
    esSecundaria,
    select,
    seg: segSmall,
    portaCheckbox,
    portaNumeroInput: inputNumero,
    portaDonanteInput: inputDonante,
    portaFields,
    lineElement: card
  };
}

// findStateSub: devuelve el objeto de estado para una subsección concreta
function findStateSub(sectionName, subName) { if (!state.sections[sectionName]) return null; return state.sections[sectionName].subsections[subName] || null; }

/* cargarOpcionesMovilDesdeCfg
   - Rellena un select según MultiPrefixes definido en structure.
   - Si se trata de una línea adicional y existe ExtraMapping, puede añadirse una opción extra condicionada.
*/
function cargarOpcionesMovilDesdeCfg(sectionName, cfg, selectEl, estado, idx, esSecundaria) {
  selectEl.innerHTML = `<option value="">-- Selecciona --</option>`;
  const mp = cfg.MultiPrefixes || '';
  const parts = mp.split(',').map(s => s.trim()).filter(Boolean);
  const map = {};
  parts.forEach(p => { const [k, pref] = p.split(':').map(x => x && x.trim()); if (k) map[k] = pref; });
  const pref = map[estado] || '';
  if (!pref) return;
  const base = catalog.filter(r => matchesPrefix(String(r.Código || ''), pref));
  base.forEach(o => selectEl.insertAdjacentHTML('beforeend', `<option value="${o.Código}">${escapeHtml(o.Plan)}</option>`));

  // Lógica para añadir opción extra según el plan principal y mapa ExtraMapping
  if (esSecundaria && estado === 'multi' && cfg.ExtraMapping) {
    const extraMap = {};
    cfg.ExtraMapping.split(';').map(x => x.trim()).filter(Boolean).forEach(p => { const [k,v] = p.split(':').map(x => x && x.trim()); if (k && v) extraMap[k] = v; });
    const stateSub = findStateSub(sectionName, cfg.Subsection);
    if (stateSub && stateSub.elementos && stateSub.elementos.lines && stateSub.elementos.lines[0]) {
      const principalSel = stateSub.elementos.lines[0].select;
      const principalVal = principalSel ? principalSel.value : '';
      if (principalVal && extraMap[principalVal]) {
        const extraPlan = catalog.find(r => r.Código === extraMap[principalVal]);
        if (extraPlan) selectEl.insertAdjacentHTML('beforeend', `<option value="${extraPlan.Código}">${escapeHtml(extraPlan.Plan)}</option>`);
      }
    }
  }

  if (!Array.from(selectEl.options).some(o => o.value === selectEl.value)) selectEl.value = '';
}

/* ------------ Recalcula detalles / facturación para una subsección móvil ------------
   - actualizarMovilSection: recopila planes seleccionados y actualiza detalles, precios y facturación.
*/
function actualizarMovilSection(sectionName, subName) {
  const stateSub = findStateSub(sectionName, subName);
  if (!stateSub) return;
  const lines = stateSub.elementos.lines || [];
  const detalles = [], precios = [];

  // Se recopilan objetos de plan por cada línea seleccionada
  lines.forEach((ln, i) => {
    const code = ln.select ? ln.select.value : '';
    const plan = findByCode(code);
    detalles.push(plan ? { name: plan.Plan, details: plan.Detalles || '' } : null);
    precios.push(plan ? { nombre: plan.Plan, promo1: plan.Promo1, meses1: plan.Meses1, valor: plan.Valor } : null);
  });

  // Se actualiza el bloque de detalles (texto)
  if (stateSub.elementos.detallesBox) {
    stateSub.elementos.detallesBox.innerHTML = '';
    detalles.forEach((d, i) => {
      const wrapper = document.createElement('div');
      wrapper.style.marginBottom = '8px';
      const hdr = document.createElement('strong');
      hdr.textContent = i === 0 ? 'Línea Principal' : `Adicional ${i}`;
      const txt = document.createElement('div');
      txt.className = 'offer-details-text';
      if (!d) {
        txt.textContent = 'Selecciona un plan para ver detalles.';
      } else {
        const safeName = escapeHtml(d.name);
        const safeDetails = escapeHtml(d.details || '');
        txt.innerHTML = `<b>${safeName}</b><br><br>${safeDetails}`;
      }
      wrapper.appendChild(hdr);
      wrapper.appendChild(txt);
      stateSub.elementos.detallesBox.appendChild(wrapper);
    });
  }

  // Se actualiza la caja de precios
  if (stateSub.elementos.preciosBox) {
    stateSub.elementos.preciosBox.innerHTML = '';
    precios.forEach((p, i) => {
      if (!p) return;
      const monthsText = (typeof p.meses1 === 'number') ? `${p.meses1} meses` : (p.meses1 || '-');
      const line = document.createElement('div');
      line.innerHTML = `Línea ${i===0 ? 'Principal' : `Adicional ${i}`}: ` +
        (p.promo1 ? `Promo1: <b>$${p.promo1}</b> (${escapeHtml(monthsText)})` : '') +
        (p.promo1 && p.valor ? ' / ' : '') +
        (p.valor ? `Sin descuento: <b>$${p.valor}</b>` : '');
      stateSub.elementos.preciosBox.appendChild(line);
    });
  }

  // Se calcula y muestra resumen de facturación
  if (stateSub.elementos.facturacionBox) {
    let totalDesc = 0, totalSin = 0;
    const rows = precios.map((p,i) => {
      if (!p) return '';
      totalDesc += Number(p.promo1 || 0);
      totalSin += Number(p.valor || 0);
      const monthsText = (typeof p.meses1 === 'number') ? `${p.meses1} meses` : (p.meses1 || '-');
      return `<div><b>${escapeHtml(p.nombre)}</b>: $${p.promo1 || '-'} (Promo1) / $${p.valor || '-'} (Sin descuento) — ${escapeHtml(monthsText)}</div>`;
    }).join('');
    stateSub.elementos.facturacionBox.innerHTML = rows + `<hr><b>Total con descuento: $${totalDesc}</b><br><b>Total sin descuento: $${totalSin}</b>`;
  }
}

/* ------------ Contrato / PDF: generación de datos y plantilla ------------
   - inicializarContrato: configura el formulario, genera .docx con docxtemplater y permite visualizar/exportar a PDF.
   - Se incluye el nombre del Ejecutivo en los datos que se pasan a la plantilla bajo la clave EJECUTIVO.
*/
function inicializarContrato() {
  // Inicialización de toggle pickup (Sucursal/Domicilio)
  (function initPickupToggleInJS() {
    const toggle = document.getElementById('pickupToggle');
    if (!toggle) return;
    toggle.querySelectorAll('button').forEach(btn => {
      btn.addEventListener('click', () => {
        toggle.querySelectorAll('button').forEach(b => {
          b.classList.remove('active');
          b.setAttribute('aria-pressed', 'false');
        });
        btn.classList.add('active');
        btn.setAttribute('aria-pressed', 'true');
      });
    });
  })();

  // Manejo del submit del formulario de contrato
  document.getElementById('contractForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    // Se recogen datos del formulario y estado actual de la UI
    const titular = document.getElementById('nombre').value || '';
    const activeSectionBtn = document.querySelector('.tab-btn.active');
    if (!activeSectionBtn) { alert('Selecciona una sección'); return; }
    const sectionName = activeSectionBtn.dataset.section;
    const activeSubName = state.sections[sectionName].activeSub;
    const stateSub = findStateSub(sectionName, activeSubName);

    // Se determina el plan principal seleccionado (si aplica)
    let planObj = null;
    if (stateSub) {
      if (stateSub.elementos && stateSub.elementos.lines) {
        for (const ln of stateSub.elementos.lines) {
          const code = ln.select ? ln.select.value : '';
          if (code) { planObj = findByCode(code); break; }
        }
      } else if (stateSub.elementos && stateSub.elementos.mainSelect) {
        planObj = findByCode(stateSub.elementos.mainSelect.value);
      }
    }

    // Generar texto MOVIL recorriendo subsecciones de Movil
    let movilParagraphs = [];

    // Para <<ALL>>: contadores por plan y totales
    const planCounts = {};
    let totalSinDescuento = 0;
    let totalConDescuento = 0;

    if (state.sections['Movil']) {
      const subsecs = state.sections['Movil'].subsections || {};
      const allLines = [];
      Object.keys(subsecs).forEach(subName => {
        const st = subsecs[subName];
        const lines = st.elementos && st.elementos.lines ? st.elementos.lines : [];
        lines.forEach(ln => { allLines.push(ln); });
      });

      // Se recorren líneas con plan seleccionado para armar párrafos y contadores
      let selectedIdx = 0;
      for (const ln of allLines) {
        const code = ln.select ? ln.select.value : '';
        const plan = findByCode(code);
        if (!plan) continue; // saltar si no hay selección

        // Contabilización para <<ALL>>
        const planName = plan.Plan || 'Sin nombre';
        planCounts[planName] = (planCounts[planName] || 0) + 1;
        totalSinDescuento += Number(plan.Valor || 0);
        totalConDescuento += Number(plan.Promo1 || plan.Valor || 0);

        // Datos de portabilidad si se marcaron
        const portaChecked = ln.portaCheckbox ? ln.portaCheckbox.checked : false;
        const numeroPorta = ln.portaNumeroInput ? (ln.portaNumeroInput.value || '').trim() : '';
        const donante = ln.portaDonanteInput ? (ln.portaDonanteInput.value || '').trim() : '';
        const valorPlan = plan.Valor || '';
        const valorPromo = plan.Promo1 || '';
        let duracion = '';
        if (plan.Meses1) {
          duracion = (typeof plan.Meses1 === 'number') ? `${plan.Meses1} meses` : `${plan.Meses1}`;
        }
        let promoPart = '';
        if (valorPromo) {
          promoPart = `y promoción de $${valorPromo}${duracion ? ' por ' + duracion : ''}`.trim();
          promoPart = promoPart.replace(/\s+$/, '');
        }

        // Construir frase base, diferenciando la primera línea seleccionada de las siguientes
        let prefix = '';
        if (selectedIdx === 0) {
          prefix = `Sr./Sra. ${titular}, Confirmamos el ${plan.Plan}, con valor normal de $${valorPlan}`;
        } else {
          prefix = `Confirmamos siguiente plan, el ${plan.Plan}, con valor normal de $${valorPlan}`;
        }

        const additions = [];
        if (promoPart) additions.push(promoPart);
        if (portaChecked && numeroPorta && donante) {
          additions.push(`portabilidad del número ${numeroPorta} desde la compañía ${donante}`);
        }
        let paragraph = prefix;
        if (additions.length > 0) {
          paragraph += ', ' + additions.join(' y ');
        }
        paragraph += '.';

        movilParagraphs.push(paragraph);
        selectedIdx++;
      }
    }

    const movilText = movilParagraphs.join('\n\n'); // Se usan dobles saltos para separar párrafos

    // Determinar si existe al menos una portabilidad seleccionada en todo Movil
    let hasAnyPortability = false;
    if (state.sections['Movil']) {
      const subsecs = state.sections['Movil'].subsections || {};
      outerLoop:
      for (const subName of Object.keys(subsecs)) {
        const st = subsecs[subName];
        const lines = st.elementos && st.elementos.lines ? st.elementos.lines : [];
        for (const ln of lines) {
          const code = ln.select ? ln.select.value : '';
          if (!code) continue;
          const portaChecked = ln.portaCheckbox ? ln.portaCheckbox.checked : false;
          if (portaChecked) { hasAnyPortability = true; break outerLoop; }
        }
      }
    }

    let condicionText = ' ';
    if (hasAnyPortability) {
      // Texto condicional que se incluye si hay alguna portabilidad
      condicionText = `¿Autoriza usted mediante esta grabación a Pacífico Cable SPA a solicitar al OAP toda información necesaria para activar el proceso? Necesito que me indique su número telefónico actual, la compañía donante, su RUT y su nombre completo.\n\nLa portabilidad solo aplica al número telefónico. Su compañía actual podría cobrar por servicios pendientes. El cambio se realiza entre 03:00 y 05:00 AM, con posible breve interrupción. En caso de retracto, puede realizarlo hasta las 20:00 horas del día en que se active el servicio.\n`;
    }

    // NOC: texto informativo según subsección activa en Movil (nuevo/cartera)
    let nocText = '';
    if (state.sections['Movil'] && state.sections['Movil'].activeSub) {
      const activeMovilSub = state.sections['Movil'].activeSub;
      if (String(activeMovilSub).toLowerCase() === 'nuevo') {
        nocText = `En Mundo, nuestros servicios tienen el cobro por mes adelantado con seis ciclos de facturación distintos con fecha de inicio 1, 5, 10, 15, 20 y 25 de cada mes. La primera boleta se emitirá en el ciclo más cercano a la activación de los servicios, con 20 días continuos de plazo para pagar. Si no se paga 5 días después, el servicio se suspende y la reposición cuesta $2.500.`;
      } else if (String(activeMovilSub).toLowerCase() === 'cartera') {
        const cicloVal = document.getElementById('ciclo') ? (document.getElementById('ciclo').value || '') : '';
        nocText = `Nuestros servicios se facturan por mes adelantado y se acoplan a su actual ciclo de facturación ${cicloVal}. Puede aplicarse un cobro proporcional el día de la activación si corresponde.`;
      } else {
        nocText = '';
      }
    }

    // OBTEN: texto según modo de pickup (Sucursal / Domicilio)
    let obtenText = '';
    const pickupToggle = document.getElementById('pickupToggle');
    let pickupMode = 'Sucursal';
    if (pickupToggle) {
      const activeBtn = pickupToggle.querySelector('button.active');
      if (activeBtn && activeBtn.dataset && activeBtn.dataset.val) pickupMode = activeBtn.dataset.val;
    }
    if (pickupMode === 'Sucursal') {
      const suc = document.getElementById('sucursal') ? (document.getElementById('sucursal').value || '') : '';
      obtenText = `En la sucursal seleccionada por usted ${suc}. El retiro y activación de su Sim Card puede realizarlo a partir del día hábil siguiente (24 horas).`;
    } else if (pickupMode === 'Domicilio') {
      const dir = document.getElementById('direccion') ? (document.getElementById('direccion').value || '') : '';
      obtenText = `La tarjeta SIM será enviada a su dirección ${dir}, en un plazo de 2 a 5 días hábiles, una vez recibida debe activarla siguiendo las indicaciones entregadas junto con su Sim Card. Si tiene dudas o consultas puede realizarlas al 6009100100 o al 442160800 opción móvil. (Activación Opción 5)`;
    }

    // ALL: resumen por plan y totales calculados
    let allText = '';
    const planEntries = Object.keys(planCounts).map(planName => {
      const count = planCounts[planName];
      const plural = count === 1 ? 'linea' : 'lineas';
      return `${count} ${plural} con el ${planName}`;
    });
    if (planEntries.length > 0) {
      const listaPlanes = planEntries.join(', ');
      allText = `Usted está contratando ${listaPlanes}, con valor total de $${totalSinDescuento} y con descuento quedaría en $${totalConDescuento}.`;
    } else {
      allText = '';
    }

    // PREPARAR DATOS Y ELEGIR PLANTILLA
    let templateFile = 'contrato_template.docx';
    const sectionNameLower = sectionName ? String(sectionName).toLowerCase() : '';
    let data = {};

    // Se incluye el nombre del Ejecutivo (si existe) en los datos para la plantilla bajo la clave EJECUTIVO.
    const ejecutivoNameForTemplate = (typeof window !== 'undefined' && window.Ejecutivo) ? String(window.Ejecutivo) : '';

    if (sectionNameLower === 'hogar') {
      // Plantilla específica para Hogar (contrato_template2.docx)
      templateFile = 'contrato_template2.docx';
      const plan = planObj || {};
      // Se calculan variantes de meses para plantillas que lo requieran
      let meses1Minus1 = '';
      const rawM1 = plan.Meses1;
      if (typeof rawM1 === 'number') {
        meses1Minus1 = rawM1 + 1;
      } else if (typeof rawM1 === 'string' && rawM1.trim() !== '' && !Number.isNaN(Number(rawM1))) {
        meses1Minus1 = Number(rawM1) + 1;
      } else {
        meses1Minus1 = '';
      }

      let meses2Plus1 = '';
      const rawM2 = plan.Meses2;
      if (typeof rawM2 === 'number') {
        meses2Plus1 = rawM2 + 1;
      } else if (typeof rawM2 === 'string' && rawM2.trim() !== '' && !Number.isNaN(Number(rawM2))) {
        meses2Plus1 = Number(rawM2) + 1;
      } else {
        meses2Plus1 = '';
      }

      data = {
        'NOMBRE': titular,
        'PLAN': plan.Plan || '',
        'DIRECCION': document.getElementById('direccion').value || '',
        'VALOR': plan.Valor || '',
        'PROMO1': plan.Promo1 || '',
        'MESES1': plan.Meses1 || '',
        'MESES1-1': meses1Minus1,
        'MESES2+1': meses2Plus1,
        'PROMO2': plan.Promo2 || '',
        'MESES2': plan.Meses2 || '',
        'DETALLES': plan.Detalles || '',
        'FECHA': document.getElementById('fecha').value || '',
        'EJECUTIVO': ejecutivoNameForTemplate
      };
    } else {
      // Plantilla por defecto (contrato_template.docx)
      templateFile = 'contrato_template.docx';
      data = {
        NOMBRE: titular,
        DIRECCION: document.getElementById('direccion').value,
        SUCURSAL: document.getElementById('sucursal').value,
        PLAN: planObj ? planObj.Plan : '',
        VALOR_PLAN: planObj ? planObj.Valor : '',
        VALOR_PROMO: planObj ? planObj.Promo1 : '',
        VALOR_PROMO2: planObj ? planObj.Promo2 : '',
        DURACION: planObj ? planObj.Meses1 : '',
        CICLO: document.getElementById('ciclo').value,
        FECHA: document.getElementById('fecha').value,
        MOVIL: movilText,
        CONDICION: condicionText,
        NOC: nocText,
        OBTEN: obtenText,
        ALL: allText,
        'EJECUTIVO': ejecutivoNameForTemplate
      };
    }

    try {
      // Se carga la plantilla, se renderiza con docxtemplater y se guarda el blob en IndexedDB.
      const content = await loadFile(templateFile);
      const zip = new PizZip(content);
      const doc = new window.docxtemplater(zip, { paragraphLoop: true, linebreaks: true, delimiters: { start: '<<', end: '>>' } });
      doc.render(data);
      const blob = doc.getZip().generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      await saveContrato(blob);
      document.getElementById('preview').innerHTML = '<p>Contrato generado y guardado. Pulsa “Visualizar contrato”.</p>';
    } catch (err) {
      // Se captura y muestra error en consola y UI.
      console.error('Error generando contrato:', err);
      showMessage('Error generando contrato. Revisa la consola.');
    }
  });

  // Visualizar y exportar a PDF el contrato guardado
  document.getElementById('visualizarButton').addEventListener('click', async () => {
    try {
      const blob = await getContrato();
      if (!blob) { alert('No hay contrato generado.'); return; }
      // Se renderiza el docx usando docx-preview en el contenedor de preview
      const archivo = new File([blob], 'Contrato.docx', { type: blob.type });
      const container = document.getElementById('preview');
      container.innerHTML = '';
      await window.docx.renderAsync(archivo, container);
      // Ajustes visuales menores para la previsualización
      const imgs = container.querySelectorAll('img'); if (imgs.length > 1) imgs[1].remove();
      const hdr = container.querySelector('div'); if (hdr) Object.assign(hdr.style, { margin: '0', padding: '0', float: 'none', display: 'block' });
      const first = container.firstElementChild; if (first) Object.assign(first.style, { margin: '0', padding: '0' });
      Object.assign(container.style, { margin: '0', padding: '0' });
      // Espera a siguiente frame para asegurar render completo
      await new Promise(requestAnimationFrame);

      // Preparación para captura a canvas y paginación con jsPDF
      const capture = document.getElementById('pdf-capture');
      const allImgs = capture.querySelectorAll('img'); allImgs.forEach(img => (img.crossOrigin = 'anonymous'));
      await Promise.all(Array.from(allImgs).map(img => new Promise(resolve => { if (img.complete) return resolve(); img.onload = resolve; img.onerror = resolve; })));
      const canvas = await html2canvas(capture, { scale: 2, useCORS: true, allowTaint: false, scrollX: 0, scrollY: -window.scrollY, width: capture.offsetWidth, height: capture.scrollHeight });
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ unit: 'mm', format: 'letter', orientation: 'portrait' });
      const pageW = pdf.internal.pageSize.getWidth(); const margin = 5;
      const pdfW = pageW - margin * 2; const pxPerMm = canvas.width / pdfW;
      const pageH = pdf.internal.pageSize.getHeight() - margin * 2; const pagePxH = Math.floor(pageH * pxPerMm);
      let renderedH = 0, pageCount = 0;
      // Se fragmenta el canvas en páginas y se añade al PDF
      while (renderedH < canvas.height) {
        const fragH = Math.min(pagePxH, canvas.height - renderedH);
        const pageCanvas = document.createElement('canvas'); pageCanvas.width = canvas.width; pageCanvas.height = fragH;
        pageCanvas.getContext('2d').drawImage(canvas, 0, renderedH, canvas.width, fragH, 0, 0, canvas.width, fragH);
        const fragImg = pageCanvas.toDataURL('image/jpeg', 1.0);
        if (pageCount > 0) pdf.addPage();
        pdf.addImage(fragImg, 'JPEG', margin, margin, pdfW, (fragH / canvas.width) * pdfW);
        renderedH += fragH; pageCount++;
      }
      // Se guarda el PDF generado (descarga)
      pdf.save('Contrato.pdf');
    } catch (err) {
      console.error('Error exportando PDF:', err);
      showMessage('Error exportando PDF. Revisa la consola.');
    }
  });
}

/* ------------ Util ------------
   - loadFile: ayuda para leer binarios usando PizZipUtils.
*/
function loadFile(url) { return new Promise((resolve, reject) => { window.PizZipUtils.getBinaryContent(url, (err, data) => err ? reject(err) : resolve(data)); }); }

/* ------------ Helpers -------------
   - escapeHtml: evita inyección HTML en textos renderizados en la UI.
*/
function escapeHtml(text) { if (text === null || text === undefined) return ''; return String(text).replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch])); }

/* ------------ Lógica de Ejecutivo: modal, persistencia, UI ------------
   - initEjecutivoUI: controla la ventana modal para editar/guardar/eliminar nombre del ejecutivo.
   - Persistencia se realiza mediante funciones saveEjecutivo/getEjecutivo/deleteEjecutivo definidas en db.js.
*/
function initEjecutivoUI() {
  const gearBtn = document.getElementById('gearBtn');
  const modal = document.getElementById('ejecutivoModal');
  const ejecutivoInput = document.getElementById('ejecutivoInput');
  const editBtn = document.getElementById('editEjecutivoBtn');
  const deleteBtn = document.getElementById('deleteEjecutivoBtn');
  const cancelBtn = document.getElementById('cancelEjecutivoBtn');
  const acceptBtn = document.getElementById('acceptEjecutivoBtn');
  const nameSpan = document.getElementById('ejecutivoName');

  // Variable global con el nombre del Ejecutivo. Se actualiza al guardar/borrar.
  window.Ejecutivo = '';

  // Carga inicial desde IndexedDB y render en header.
  (async function loadAndRender() {
    try {
      const stored = (typeof getEjecutivo === 'function') ? await getEjecutivo() : '';
      window.Ejecutivo = stored || '';
      renderEjecutivoName();
    } catch (err) {
      console.error('No se pudo leer Ejecutivo desde IndexedDB', err);
    }
  })();

  // renderEjecutivoName: muestra/oculta nombre en el header.
  function renderEjecutivoName() {
    if (!nameSpan) return;
    if (window.Ejecutivo && String(window.Ejecutivo).trim() !== '') {
      nameSpan.textContent = String(window.Ejecutivo);
      nameSpan.title = `Ejecutivo: ${window.Ejecutivo}`;
    } else {
      nameSpan.textContent = '';
      nameSpan.title = '';
    }
  }

  // openModal / closeModal: control visual del diálogo
  function openModal() {
    if (!modal) return;
    ejecutivoInput.value = window.Ejecutivo || '';
    ejecutivoInput.setAttribute('readonly', 'readonly');
    if (!window.Ejecutivo) {
      // Si no hay nombre guardado, se permite escribir de inmediato.
      ejecutivoInput.removeAttribute('readonly');
      ejecutivoInput.focus();
    }
    modal.hidden = false;
    acceptBtn.focus();
  }

  function closeModal() {
    if (!modal) return;
    modal.hidden = true;
  }

  // Evento para abrir modal desde botón engranaje
  gearBtn && gearBtn.addEventListener('click', (e) => {
    openModal();
  });

  // Botón editar: habilita el input para modificar
  editBtn && editBtn.addEventListener('click', () => {
    ejecutivoInput.removeAttribute('readonly');
    ejecutivoInput.focus();
    const val = ejecutivoInput.value;
    ejecutivoInput.value = '';
    ejecutivoInput.value = val;
  });

  // Botón eliminar: limpia el input para borrar y guardar vacío si se acepta
  deleteBtn && deleteBtn.addEventListener('click', async () => {
    ejecutivoInput.removeAttribute('readonly');
    ejecutivoInput.value = '';
    ejecutivoInput.focus();
  });

  // Cancel: restaura valor previo y cierra modal
  cancelBtn && cancelBtn.addEventListener('click', (e) => {
    e.preventDefault();
    ejecutivoInput.value = window.Ejecutivo || '';
    ejecutivoInput.setAttribute('readonly', 'readonly');
    closeModal();
  });

  // Accept: guarda o borra el nombre usando las funciones definidas en db.js
  acceptBtn && acceptBtn.addEventListener('click', async (e) => {
    e.preventDefault();
    const newName = (ejecutivoInput.value || '').trim();
    try {
      if (!newName) {
        if (typeof deleteEjecutivo === 'function') await deleteEjecutivo();
        window.Ejecutivo = '';
      } else {
        if (typeof saveEjecutivo === 'function') await saveEjecutivo(newName);
        window.Ejecutivo = newName;
      }
      renderEjecutivoName();
      ejecutivoInput.setAttribute('readonly', 'readonly');
      closeModal();
    } catch (err) {
      console.error('Error guardando/eliminando Ejecutivo', err);
      showMessage('No se pudo guardar el nombre del Ejecutivo. Revisa la consola.', true, 6000);
    }
  });

  // Cerrar modal con ESC
  document.addEventListener('keydown', (ev) => {
    if (ev.key === 'Escape') {
      const modalVisible = modal && !modal.hidden;
      if (modalVisible) {
        cancelBtn && cancelBtn.click();
      }
    }
  });

  // Cerrar modal si se hace click fuera del diálogo
  modal && modal.addEventListener('click', (ev) => {
    if (ev.target === modal) {
      cancelBtn && cancelBtn.click();
    }
  });
}