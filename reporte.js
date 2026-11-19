'use strict';

const byId = (id) => document.getElementById(id);
const elements = Object.fromEntries(['file', 'download', 'file-name', 'warehouse', 'status', 'status-label', 'status-menu', 'order-type', 'order-presence', 'request-type', 'search', 'clear', 'count', 'summary-columns', 'detail-columns', 'table-wrap', 'table-head', 'table-body', 'message', 'save-status', 'storage-status'].map((id) => [id, byId(id)]));
const state = { headers: [], rows: [], summary: new Set(), detail: new Set(), expanded: new Set(), statuses: new Set(), keys: {} };
const normalize = (value) => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
const escapeHtml = (value) => String(value ?? '').replace(/[&<>"']/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[character]);
const valueOf = (row, key) => String(row[key] ?? '').trim();
const storageKey = 'seguimiento-compras-v1';
const pendingKey = 'seguimiento-compras-pendiente-v1';
const columnsKey = 'seguimiento-compras-columnas-v1';
let savedEntries = {};
let pendingEntries = {};
let columnPreferences = null;
try { savedEntries = JSON.parse(localStorage.getItem(storageKey) || '{}'); } catch { savedEntries = {}; }
try { pendingEntries = JSON.parse(localStorage.getItem(pendingKey) || '{}'); } catch { pendingEntries = {}; }
try {
  const saved = JSON.parse(localStorage.getItem(columnsKey) || 'null');
  if (saved && Array.isArray(saved.summary) && Array.isArray(saved.detail) && saved.summary.every((value) => typeof value === 'string') && saved.detail.every((value) => typeof value === 'string')) columnPreferences = saved;
} catch { columnPreferences = null; }
let databaseAvailable = false;
let writeQueue = Promise.resolve();
const documentNumber = (row) => {
  const document = valueOf(row, state.keys.order);
  return document.match(/\d+\s*$/)?.[0].trim() ?? document;
};
const uniqueValues = (rows, key) => [...new Set(rows.map((row) => valueOf(row, key)).filter(Boolean))].sort((first, second) => first.localeCompare(second, 'es', { numeric: true }));
const groupKey = (row) => JSON.stringify([documentNumber(row), valueOf(row, state.keys.provider)]);

function saveEntries(changes) {
  for (const [key, values] of Object.entries(changes)) pendingEntries[key] = { ...pendingEntries[key], ...values };
  try {
    localStorage.setItem(storageKey, JSON.stringify(savedEntries));
    localStorage.setItem(pendingKey, JSON.stringify(pendingEntries));
  } catch { elements['save-status'].textContent = 'No se pudo guardar una copia temporal en el navegador.'; }
  if (!databaseAvailable) {
    elements['storage-status'].textContent = 'Sin conexión a la base de COMPRAS. Inicia el reporte con iniciar-reporte.cmd; descarga el Excel para conservar los cambios.';
    return;
  }
  elements['storage-status'].textContent = 'Guardando en la base de COMPRAS...';
  const operation = writeQueue.then(async () => {
    const snapshot = JSON.parse(JSON.stringify(pendingEntries));
    if (!Object.keys(snapshot).length) return;
    const response = await fetch('/api/seguimiento', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(snapshot) });
    if (!response.ok) throw new Error('No se pudo guardar');
    for (const [key, values] of Object.entries(snapshot)) {
      for (const [field, value] of Object.entries(values)) {
        if (pendingEntries[key]?.[field] === value) delete pendingEntries[key][field];
      }
      if (!Object.keys(pendingEntries[key] || {}).length) delete pendingEntries[key];
    }
    try { localStorage.setItem(pendingKey, JSON.stringify(pendingEntries)); } catch { }
  });
  writeQueue = operation.catch(() => {});
  operation.then(() => {
    elements['storage-status'].textContent = Object.keys(pendingEntries).length ? 'Guardando en la base de COMPRAS...' : 'Seguimiento guardado en la carpeta COMPRAS.';
  }).catch(() => {
    elements['storage-status'].textContent = 'No se pudo guardar en la base de COMPRAS. Descarga el Excel y revisa que el servidor esté abierto.';
  });
}

function dateValue(cell) {
  if (!cell || cell.v == null || cell.v === '') return '';
  if (typeof cell.v === 'number') {
    const date = XLSX.SSF.parse_date_code(cell.v);
    return date ? `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}` : '';
  }
  if (cell.v instanceof Date) return `${cell.v.getFullYear()}-${String(cell.v.getMonth() + 1).padStart(2, '0')}-${String(cell.v.getDate()).padStart(2, '0')}`;
  const text = String(cell.v).trim();
  const iso = text.match(/^(\d{4}-\d{2}-\d{2})/);
  if (iso) return iso[1];
  const parts = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
  return parts ? `${parts[3]}-${parts[2].padStart(2, '0')}-${parts[1].padStart(2, '0')}` : '';
}

function inspectWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, blankrows: true });
    const headerIndex = data.findIndex((row) => findHeader(row, ['documentodecompra']) && findHeader(row, ['almacenpedido']) && findHeader(row, ['estadooc']) && findHeader(row, ['proveedor']) && findHeader(row, ['codigo']));
    if (headerIndex >= 0) return { workbook, sheet, sheetName, data, headerIndex, range: XLSX.utils.decode_range(sheet['!ref']) };
  }
  throw new Error('No se encontró una hoja con las columnas Documento de Compra, Proveedor, Código, Almacén Pedido y Estado O.C.');
}

function importEntries(source, overwrite) {
  const headers = source.data[source.headerIndex].map((header) => String(header).trim());
  const orderIndex = headers.findIndex((header) => normalize(header) === 'documentodecompra');
  const providerIndex = headers.findIndex((header) => normalize(header) === 'proveedor');
  const etaIndex = headers.findIndex((header) => normalize(header) === 'eta');
  const destinationIndex = headers.findIndex((header) => normalize(header) === 'almacendestino');
  if (etaIndex < 0 && destinationIndex < 0) return 0;
  let imported = 0;
  const changes = {};
  for (let index = source.headerIndex + 1; index < source.data.length; index++) {
    const row = source.data[index];
    const document = String(row[orderIndex] ?? '').trim();
    if (!document) continue;
    const order = document.match(/\d+\s*$/)?.[0].trim() ?? document;
    const key = JSON.stringify([order, String(row[providerIndex] ?? '').trim()]);
    const dateCell = etaIndex < 0 ? null : source.sheet[XLSX.utils.encode_cell({ r: source.range.s.r + index, c: source.range.s.c + etaIndex })];
    const values = { eta: dateValue(dateCell), destination: destinationIndex < 0 ? '' : String(row[destinationIndex] ?? '').trim() };
    for (const field of ['eta', 'destination']) {
      if (values[field] && (overwrite || !Object.hasOwn(savedEntries[key] || {}, field))) {
        savedEntries[key] ||= {};
        savedEntries[key][field] = values[field];
        changes[key] ||= {};
        changes[key][field] = values[field];
        imported++;
      }
    }
  }
  if (imported) saveEntries(changes);
  return imported;
}

function showMessage(title, text) {
  elements.message.innerHTML = `<div class="empty-icon">↗</div><h3>${escapeHtml(title)}</h3><p>${escapeHtml(text)}</p>`;
  elements.message.hidden = false;
  elements['table-wrap'].hidden = true;
}

function populateSelect(element, values, placeholder, previous = '') {
  element.innerHTML = `<option value="">${escapeHtml(placeholder)}</option>` + values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join('');
  element.value = values.includes(previous) ? previous : '';
}

function findHeader(headers, names) {
  return headers.find((header) => names.includes(normalize(header)));
}

function updateSubcontractStatuses(rows, source, headers) {
  const purchaseColumn = headers.findIndex((header) => normalize(header) === 'cantidaddecompra');
  const guideColumn = headers.findIndex((header) => normalize(header) === 'cantidaddelaguia');
  if (purchaseColumn < 0 || guideColumn < 0) throw new Error('Faltan las columnas Cantidad de Compra o Cantidad de la Guía.');
  const orderGroups = new Map();
  for (const row of rows) {
    const key = groupKey(row);
    if (!orderGroups.has(key)) orderGroups.set(key, []);
    orderGroups.get(key).push(row);
  }
  const statuses = new Map();
  for (const [key, group] of orderGroups) {
    if (!group.some((row) => normalize(valueOf(row, state.keys.status)).includes('subcontrato'))) continue;
    const items = new Map();
    for (const row of group) {
      const readQuantity = (column, defaultValue) => {
        const cell = source.sheet[XLSX.utils.encode_cell({ r: row.sourceRow, c: source.range.s.c + column })];
        if (!cell || cell.v === '' || cell.v == null) return defaultValue;
        const number = typeof cell.v === 'number' ? cell.v : Number(String(cell.v).trim().replace(',', '.'));
        return Number.isFinite(number) ? number : NaN;
      };
      const purchase = readQuantity(purchaseColumn, NaN);
      const guide = readQuantity(guideColumn, 0);
      if (!Number.isFinite(purchase) || !Number.isFinite(guide) || purchase < 0 || guide < 0) throw new Error(`Cantidades inválidas en la orden ${documentNumber(row)}. Revisa el Excel.`);
      const itemKey = JSON.stringify([valueOf(row, state.keys.request), valueOf(row, state.keys.code), valueOf(row, state.keys.item), purchase]);
      const item = items.get(itemKey) || { purchase, guide: 0 };
      item.guide += guide;
      items.set(itemKey, item);
    }
    const quantities = [...items.values()];
    const status = quantities.every((item) => item.guide === 0) ? 'Aprobado' : quantities.every((item) => Math.abs(item.guide - item.purchase) < 1e-8) ? 'En Almacén' : 'Parcialmente en Almacén';
    statuses.set(key, status);
    for (const row of group) if (normalize(valueOf(row, state.keys.status)).includes('subcontrato')) row[state.keys.status] = status;
  }
  return statuses;
}

function loadWorkbook(buffer, name) {
  if (typeof XLSX === 'undefined') throw new Error('No se pudo cargar el lector de Excel. Comprueba que el archivo xlsx.full.min.js esté junto al reporte.');
  const selected = inspectWorkbook(buffer);
  const headers = selected.data[selected.headerIndex].map((header) => String(header).trim());
  if (new Set(headers.filter(Boolean)).size !== headers.filter(Boolean).length) throw new Error('La hoja contiene nombres de columnas repetidos. Renómbralos en Excel antes de cargarla.');
  state.headers = headers.filter(Boolean);
  state.keys = { order: findHeader(headers, ['documentodecompra']), provider: findHeader(headers, ['proveedor']), warehouse: findHeader(headers, ['almacenpedido']), status: findHeader(headers, ['estadooc']), code: findHeader(headers, ['codigo']), request: findHeader(headers, ['pedido']), requestType: findHeader(headers, ['pedidoaalmacencompras']), item: findHeader(headers, ['correlativo']) };
  state.source = selected;
  state.fileName = name;
  importEntries(selected, false);
  state.rows = selected.data.slice(selected.headerIndex + 1).map((row, index) => ({ ...Object.fromEntries(headers.map((header, column) => [header, row[column] ?? '']).filter(([header]) => header)), sourceRow: selected.range.s.r + selected.headerIndex + 1 + index })).filter((row) => headers.some((header) => header && valueOf(row, header)));
  state.subcontractStatuses = updateSubcontractStatuses(state.rows.filter((row) => valueOf(row, state.keys.order)), selected, headers);
  const typeHeader = findHeader(state.headers, ['tipodeoc']) || 'Tipo de OC';
  if (!state.headers.includes(typeHeader)) state.headers.push(typeHeader);
  for (const [normalized, title] of [['eta', 'ETA'], ['almacendestino', 'Almacén destino']]) {
    if (!findHeader(state.headers, [normalized])) state.headers.push(title);
  }
  state.keys.type = typeHeader;
  for (const row of state.rows) row[typeHeader] = classifyCode(valueOf(row, state.keys.code));
  if (!columnPreferences) columnPreferences = { summary: [normalize(typeHeader)], detail: ['pedido', 'codigo', 'recurso', 'unidad', 'cantidadpedida', 'cantidaddecompra', 'cantidaddelaguia', 'tipodeoc'] };
  state.summary = new Set(state.headers.filter((header) => columnPreferences.summary.includes(normalize(header))));
  state.detail = new Set(state.headers.filter((header) => columnPreferences.detail.includes(normalize(header))));
  state.expanded.clear();
  elements['file-name'].textContent = name;
  elements.search.value = '';
  elements['order-type'].value = '';
  elements['order-presence'].value = '';
  elements['request-type'].value = '';
  populateSelect(elements.warehouse, uniqueValues(state.rows, state.keys.warehouse), 'Todos los almacenes');
  state.statuses.clear();
  elements.status.innerHTML = uniqueValues(state.rows, state.keys.status).map((value) => `<label><input type="checkbox" value="${escapeHtml(value)}">${escapeHtml(value)}</label>`).join('');
  elements['status-label'].textContent = 'Todos los estados';
  for (const element of [elements.warehouse, elements['order-type'], elements['order-presence'], elements.search, elements.clear]) element.disabled = false;
  elements['request-type'].disabled = !state.keys.requestType;
  elements.download.disabled = false;
  renderColumnChoices();
  render();
}

function renderColumnChoices() {
  const summaryHeaders = state.headers.filter((header) => header !== state.keys.order && header !== state.keys.provider && !['eta', 'almacendestino'].includes(normalize(header)));
  for (const [container, headers, selected, type] of [[elements['summary-columns'], summaryHeaders, state.summary, 'summary'], [elements['detail-columns'], state.headers, state.detail, 'detail']]) {
    container.innerHTML = headers.map((header) => `<label><input type="checkbox" data-type="${type}" value="${escapeHtml(header)}" ${selected.has(header) ? 'checked' : ''}><span>${escapeHtml(header)}</span></label>`).join('');
  }
}

function filteredRows() {
  const query = elements.search.value.trim().toLocaleLowerCase('es');
  return state.rows.filter((row) => (!elements.warehouse.value || valueOf(row, state.keys.warehouse) === elements.warehouse.value) && (!state.statuses.size || state.statuses.has(valueOf(row, state.keys.status))) && (!elements['order-type'].value || valueOf(row, state.keys.type) === elements['order-type'].value) && (!elements['order-presence'].value || (elements['order-presence'].value === 'con') === Boolean(valueOf(row, state.keys.order))) && (!elements['request-type'].value || normalize(valueOf(row, state.keys.requestType)).startsWith(`pedidoa${elements['request-type'].value}`)) && (!query || [documentNumber(row) || 'Sin OC', valueOf(row, state.keys.provider), valueOf(row, state.keys.requestType), valueOf(row, state.keys.request)].some((value) => value.toLocaleLowerCase('es').includes(query))));
}

function displayGroupValue(rows, header) {
  const values = uniqueValues(rows, header);
  return values.length ? values.join(' / ') : '—';
}

function render() {
  const rows = filteredRows();
  const groups = new Map();
  for (const row of rows) {
    const key = valueOf(row, state.keys.order) ? groupKey(row) : JSON.stringify(['sin-documento']);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }
  const orderedGroups = [...groups.entries()].sort(([first], [second]) => first.localeCompare(second, 'es', { numeric: true }));
  const groupLabel = groups.size === 1 ? 'grupo' : 'grupos';
  elements.count.textContent = `${groups.size} ${groupLabel} · ${rows.length} ${rows.length === 1 ? 'registro' : 'registros'}`;
  if (!rows.length) { showMessage('Sin resultados', 'No hay registros para los filtros seleccionados.'); return; }
  const summaryHeaders = state.headers.filter((header) => state.summary.has(header) && !['eta', 'almacendestino'].includes(normalize(header)));
  const detailHeaders = state.headers.filter((header) => state.detail.has(header));
  const allHeaders = [state.keys.order, state.keys.provider, ...summaryHeaders, 'ETA', 'Almacén destino'];
  elements['table-head'].innerHTML = `<tr>${allHeaders.map((header) => `<th scope="col">${escapeHtml(header)}</th>`).join('')}</tr>`;
  elements['table-body'].innerHTML = orderedGroups.map(([key, group], index) => {
    const order = documentNumber(group[0]);
    const expanded = state.expanded.has(key);
    const saved = order ? savedEntries[key] || {} : {};
    const cells = [state.keys.provider, ...summaryHeaders].map((header) => `<td>${escapeHtml(displayGroupValue(group, header))}</td>`).join('');
    const detail = expanded ? `<tr><td class="detail-cell" colspan="${allHeaders.length}"><div class="detail-header">Detalle de pedidos <span>${group.length} ${group.length === 1 ? 'registro' : 'registros'}</span></div>${detailHeaders.length ? `<div class="detail-scroll"><table class="detail-table"><thead><tr>${detailHeaders.map((header) => `<th scope="col">${escapeHtml(header)}</th>`).join('')}</tr></thead><tbody>${group.map((row) => `<tr>${detailHeaders.map((header) => `<td>${escapeHtml((header === state.keys.order ? documentNumber(row) || 'Sin OC' : normalize(header) === 'eta' ? saved.eta : normalize(header) === 'almacendestino' ? saved.destination : valueOf(row, header)) || '—')}</td>`).join('')}</tr>`).join('')}</tbody></table></div>` : '<p>Selecciona columnas para mostrar el detalle.</p>'}</td></tr>` : '';
    return `<tr class="order-row"><td><button class="order-button" data-index="${index}" aria-expanded="${expanded}"><span>${expanded ? '▾' : '▸'}</span>${escapeHtml(order || 'Sin OC')}</button></td>${cells}<td class="editable-cell"><input type="date" aria-label="ETA ${escapeHtml(order)}" data-index="${index}" data-field="eta" value="${escapeHtml(saved.eta || '')}" ${order ? '' : 'disabled'}></td><td class="editable-cell"><input type="text" aria-label="Almacén destino ${escapeHtml(order)}" data-index="${index}" data-field="destination" value="${escapeHtml(saved.destination || '')}" ${order ? '' : 'disabled'}></td></tr>${detail}`;
  }).join('');
  elements['table-body'].querySelectorAll('.order-button').forEach((button) => button.addEventListener('click', () => {
    const order = orderedGroups[Number(button.dataset.index)][0];
    if (state.expanded.has(order)) state.expanded.delete(order); else state.expanded.add(order);
    render();
  }));
  elements['table-body'].querySelectorAll('.editable-cell input').forEach((input) => input.addEventListener('change', () => {
    const key = orderedGroups[Number(input.dataset.index)][0];
    savedEntries[key] ||= {};
    savedEntries[key][input.dataset.field] = input.value.trim();
    saveEntries({ [key]: { [input.dataset.field]: savedEntries[key][input.dataset.field] } });
  }));
  elements.message.hidden = true;
  elements['table-wrap'].hidden = false;
}

function classifyCode(code) {
  return code ? (/^(02A|10)/i.test(code) ? 'COMPRA' : 'SERVICIO') : '';
}

function exportWorkbook() {
  const source = state.source;
  const headers = source.data[source.headerIndex].map((header) => String(header).trim());
  const columns = {};
  for (const [field, title] of [['type', 'Tipo de OC'], ['eta', 'ETA'], ['destination', 'Almacén destino']]) {
    let index = headers.findIndex((header) => normalize(header) === normalize(title));
    if (index < 0) { index = headers.length; headers.push(title); }
    columns[field] = source.range.s.c + index;
    source.sheet[XLSX.utils.encode_cell({ r: source.range.s.r + source.headerIndex, c: columns[field] })] = { t: 's', v: title };
  }
  for (let index = source.headerIndex + 1; index < source.data.length; index++) {
    const values = source.data[index];
    if (!values.some((value) => String(value ?? '').trim())) continue;
    const row = Object.fromEntries(headers.map((header, column) => [header, values[column] ?? '']).filter(([header]) => header));
    const rowNumber = source.range.s.r + index;
    const type = classifyCode(valueOf(row, state.keys.code));
    const saved = savedEntries[groupKey(row)] || {};
    const cell = (field) => XLSX.utils.encode_cell({ r: rowNumber, c: columns[field] });
    if (type) source.sheet[cell('type')] = { t: 's', v: type }; else delete source.sheet[cell('type')];
    if (!valueOf(row, state.keys.order)) continue;
    const status = state.subcontractStatuses.get(groupKey(row));
    if (status && normalize(valueOf(row, state.keys.status)).includes('subcontrato')) source.sheet[XLSX.utils.encode_cell({ r: rowNumber, c: source.range.s.c + headers.indexOf(state.keys.status) })] = { t: 's', v: status };
    if (saved.eta && /^\d{4}-\d{2}-\d{2}$/.test(saved.eta)) {
      const serial = (Date.parse(`${saved.eta}T00:00:00Z`) - Date.UTC(1899, 11, 30)) / 86400000;
      source.sheet[cell('eta')] = { t: 'n', v: serial, z: 'dd/mm/yyyy' };
    } else delete source.sheet[cell('eta')];
    if (saved.destination) source.sheet[cell('destination')] = { t: 's', v: saved.destination }; else delete source.sheet[cell('destination')];
  }
  source.range.e.c = Math.max(source.range.e.c, ...Object.values(columns));
  source.sheet['!ref'] = XLSX.utils.encode_range(source.range);
  XLSX.writeFile(source.workbook, `${state.fileName.replace(/\.[^.]+$/, '')}_seguimiento.xlsx`, { bookType: 'xlsx' });
  elements['save-status'].textContent = 'Excel de seguimiento descargado.';
}

elements.file.addEventListener('change', async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    if (databaseAvailable) {
      await writeQueue;
      if (Object.keys(pendingEntries).length) {
        saveEntries({});
        await writeQueue;
        if (Object.keys(pendingEntries).length) throw new Error('Hay cambios pendientes que no se pudieron guardar en la base. Descarga una copia y revisa el servidor.');
      }
      const response = await fetch('/api/seguimiento', { cache: 'no-store' });
      if (!response.ok) throw new Error('No se pudo consultar la base de seguimiento. Revisa que el servidor esté abierto.');
      const entries = await response.json();
      savedEntries = Object.fromEntries([...new Set([...Object.keys(savedEntries), ...Object.keys(entries)])].map((key) => [key, { ...savedEntries[key], ...entries[key] }]));
    }
    loadWorkbook(await file.arrayBuffer(), file.name);
  }
  catch (error) { showMessage('No se pudo abrir el archivo', error.message || 'Comprueba que el archivo sea un Excel válido.'); }
  finally { event.target.value = ''; }
});
elements.download.addEventListener('click', () => {
  try { exportWorkbook(); }
  catch (error) { elements['save-status'].textContent = `No se pudo descargar el Excel: ${error.message}`; }
});
for (const element of [elements.warehouse, elements['order-type'], elements['order-presence'], elements['request-type'], elements.search]) element.addEventListener('input', render);
elements.status.addEventListener('change', (event) => {
  const checkbox = event.target;
  if (!checkbox.matches('input[type="checkbox"]')) return;
  if (checkbox.checked) state.statuses.add(checkbox.value); else state.statuses.delete(checkbox.value);
  elements['status-label'].textContent = state.statuses.size ? `${state.statuses.size} estado(s) seleccionado(s)` : 'Todos los estados';
  render();
});
elements.clear.addEventListener('click', () => { elements.warehouse.value = ''; elements['order-type'].value = ''; elements['order-presence'].value = ''; elements['request-type'].value = ''; state.statuses.clear(); elements.status.querySelectorAll('input').forEach((input) => { input.checked = false; }); elements['status-label'].textContent = 'Todos los estados'; elements.search.value = ''; render(); });
for (const container of [elements['summary-columns'], elements['detail-columns']]) container.addEventListener('change', (event) => {
  const checkbox = event.target;
  if (!checkbox.matches('input[type="checkbox"]')) return;
  const selected = state[checkbox.dataset.type];
  if (checkbox.checked) selected.add(checkbox.value); else selected.delete(checkbox.value);
  const preferences = new Set(columnPreferences[checkbox.dataset.type]);
  if (checkbox.checked) preferences.add(normalize(checkbox.value)); else preferences.delete(normalize(checkbox.value));
  columnPreferences[checkbox.dataset.type] = [...preferences];
  try { localStorage.setItem(columnsKey, JSON.stringify(columnPreferences)); }
  catch { elements['save-status'].textContent = 'No se pudo guardar la selección de columnas en el navegador.'; }
  render();
});

async function connectDatabase() {
  try {
    if (location.protocol === 'file:') throw new Error('El reporte se abrió directamente');
    const response = await fetch('/api/seguimiento', { cache: 'no-store' });
    if (!response.ok) throw new Error('Base no disponible');
    const entries = await response.json();
    const missing = {};
    for (const [key, values] of Object.entries(savedEntries)) {
      for (const field of ['eta', 'destination']) {
        if (typeof values?.[field] === 'string' && !Object.hasOwn(entries[key] || {}, field)) {
          missing[key] ||= {};
          missing[key][field] = values[field];
        }
      }
    }
    savedEntries = Object.fromEntries([...new Set([...Object.keys(savedEntries), ...Object.keys(entries), ...Object.keys(pendingEntries)])].map((key) => [key, { ...savedEntries[key], ...entries[key], ...pendingEntries[key] }]));
    databaseAvailable = true;
    elements['storage-status'].textContent = 'Base de seguimiento conectada: carpeta COMPRAS.';
    if (Object.keys(missing).length || Object.keys(pendingEntries).length) saveEntries(missing);
  } catch {
    elements['storage-status'].textContent = 'Sin conexión a la base de COMPRAS. Abre iniciar-reporte.cmd para guardar en la carpeta.';
  } finally {
    elements.file.disabled = false;
  }
}

connectDatabase();