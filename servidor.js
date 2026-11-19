'use strict';

const http = require('http');
const fs = require('fs/promises');
const path = require('path');
const { execFile } = require('child_process');

const host = '127.0.0.1';
const port = Number(process.env.PUERTO_REPORTE || 3783);
const database = process.env.BASE_SEGUIMIENTO || path.join(__dirname, 'seguimiento-datos.json');
const files = { '/': ['reporte.html', 'text/html; charset=utf-8'], '/reporte.html': ['reporte.html', 'text/html; charset=utf-8'], '/reporte.css': ['reporte.css', 'text/css; charset=utf-8'], '/reporte.js': ['reporte.js', 'text/javascript; charset=utf-8'], '/xlsx.full.min.js': ['xlsx.full.min.js', 'text/javascript; charset=utf-8'] };
let writes = Promise.resolve();

async function readDatabase() {
  try {
    const data = JSON.parse(await fs.readFile(database, 'utf8'));
    if (data.version !== 1 || !data.registros || typeof data.registros !== 'object' || Array.isArray(data.registros)) throw new Error('Formato de base de datos inválido');
    return data;
  } catch (error) {
    if (error.code === 'ENOENT') return { version: 1, registros: {} };
    throw error;
  }
}

function send(response, status, data) {
  response.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
  response.end(JSON.stringify(data));
}

async function updateDatabase(changes) {
  const data = await readDatabase();
  for (const [key, values] of Object.entries(changes)) {
    let identifiers;
    try { identifiers = JSON.parse(key); } catch { throw new Error('Clave de seguimiento inválida'); }
    if (!Array.isArray(identifiers) || identifiers.length !== 2 || !identifiers.every((value) => typeof value === 'string') || !identifiers[0] || !values || typeof values !== 'object' || Array.isArray(values)) throw new Error('Registro de seguimiento inválido');
    const current = data.registros[key] || {};
    for (const field of ['eta', 'destination']) {
      if (Object.hasOwn(values, field)) {
        if (typeof values[field] !== 'string' || values[field].length > 500 || (field === 'eta' && values[field] && !/^\d{4}-\d{2}-\d{2}$/.test(values[field]))) throw new Error('Valor de seguimiento inválido');
        current[field] = values[field];
      }
    }
    data.registros[key] = current;
  }
  const temporary = `${database}.tmp`;
  await fs.writeFile(temporary, JSON.stringify(data, null, 2), 'utf8');
  await fs.rename(temporary, database);
  return data.registros;
}

const server = http.createServer(async (request, response) => {
  const allowedHosts = [`${host}:${port}`, `localhost:${port}`];
  if (!allowedHosts.includes(request.headers.host)) return send(response, 403, { error: 'Acceso no autorizado' });
  const origin = request.headers.origin;
  if (origin && !allowedHosts.some((allowed) => origin === `http://${allowed}`)) return send(response, 403, { error: 'Origen no autorizado' });
  try {
    const url = new URL(request.url, `http://${request.headers.host}`);
    if (request.method === 'GET' && url.pathname === '/api/seguimiento') return send(response, 200, (await readDatabase()).registros);
    if (request.method === 'POST' && url.pathname === '/api/seguimiento') {
      if (!origin || !request.headers['content-type']?.startsWith('application/json')) return send(response, 403, { error: 'Solicitud no autorizada' });
      let body = '';
      for await (const chunk of request) {
        body += chunk;
        if (body.length > 5_000_000) return send(response, 413, { error: 'Solicitud demasiado grande' });
      }
      const changes = JSON.parse(body);
      if (!changes || typeof changes !== 'object' || Array.isArray(changes)) return send(response, 400, { error: 'Datos inválidos' });
      const operation = writes.then(() => updateDatabase(changes));
      writes = operation.catch(() => {});
      return send(response, 200, await operation);
    }
    if (request.method === 'GET' && files[url.pathname]) {
      const [name, type] = files[url.pathname];
      const content = await fs.readFile(path.join(__dirname, name));
      response.writeHead(200, { 'Content-Type': type, 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
      return response.end(content);
    }
    return send(response, 404, { error: 'No encontrado' });
  } catch (error) {
    console.error(error.message);
    return send(response, 500, { error: 'No se pudo acceder al archivo de seguimiento. Revise la consola del servidor.' });
  }
});

server.on('error', (error) => {
  console.error(`No se pudo iniciar el reporte: ${error.message}`);
  if (error.code === 'EADDRINUSE') console.error('El puerto 3783 ya está en uso. Si el reporte ya está abierto, visite http://127.0.0.1:3783/');
});
server.listen(port, host, () => {
  console.log(`Reporte disponible en http://${host}:${port}/`);
  console.log(`Base de seguimiento: ${database}`);
  if (process.env.ABRIR_REPORTE === '1' && process.platform === 'win32') execFile('cmd.exe', ['/c', 'start', '', `http://${host}:${port}/`], { windowsHide: true });
});