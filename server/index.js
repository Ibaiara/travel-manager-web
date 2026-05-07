const express = require('express');
const fs = require('fs');          // sync helpers (existsSync, mkdirSync)
const fsp = fs.promises;           // async (readFile, writeFile, access...)
const path = require('path');
const cors = require('cors');

const XLSX = require('xlsx');
const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage() });



// TODO: pon aquí el nombre exacto de columna (header) de cada Excel
const CLIENTES_COLUMN = 'Account Name';      // <-- cámbialo
const PROYECTOS_COLUMNS = [ 'Project Name / Reference', 'Reference',  'Topic'];

const app = express();
app.use(express.json({ limit: '2mb' }));
app.use(cors());

// ============================================
// CONFIGURACIÓN PARA MODO PORTABLE (.exe)
// ============================================

const rutaBase =__dirname;

// Configuración de rutas (funciona tanto en desarrollo como en .exe)
const DATA_DIR = process.env.DATA_DIR || path.join(rutaBase, 'data');
const DATA_FILE = path.join(DATA_DIR, 'viajes_data.json');
const BACKUP_DIR = path.join(DATA_DIR, 'backups');
const LOCK_FILE = path.join(DATA_DIR, '.lock');

const PROYECTOS_FILE = path.join(DATA_DIR, 'ofertas.json');

const LOCATIONS_FILE = path.join(DATA_DIR, 'locations.json');

const USERS_FILE = path.join(DATA_DIR, 'users.json');

// Catálogo clientes (para autocompletar) - recomendado: DATA_DIR/clientes.json
const CLIENTES_FILE = path.join(DATA_DIR, 'clientes.json');

const OFERTAS_FILE = path.join(DATA_DIR, 'ofertas.json');

// Fallback coords (HQ)
const office = [43.04426527618791, -2.2100984760320834];

// Crear directorios si no existen
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  console.log('✅ Carpeta de datos creada:', DATA_DIR);
}
if (!fs.existsSync(BACKUP_DIR)) {
  fs.mkdirSync(BACKUP_DIR, { recursive: true });
  console.log('✅ Carpeta de backups creada:', BACKUP_DIR);
}

// ============================================
// SISTEMA DE LOCKS PARA CONCURRENCIA
// ============================================
class FileLock {
  constructor(lockFile) {
    this.lockFile = lockFile;
    this.maxRetries = 10;
    this.retryDelay = 100;
  }

  async acquire() {
    for (let i = 0; i < this.maxRetries; i++) {
      try {
        await fsp.writeFile(this.lockFile, process.pid.toString(), { flag: 'wx' });
        return true;
      } catch (error) {
        if (error.code === 'EEXIST') {
          await new Promise(resolve => setTimeout(resolve, this.retryDelay));
        } else {
          throw error;
        }
      }
    }
    throw new Error('No se pudo adquirir el lock después de varios intentos');
  }

  async release() {
    try {
      await fsp.unlink(this.lockFile);
    } catch {
      // Ignorar
    }
  }
}
const fileLock = new FileLock(LOCK_FILE);

// ============================================
// FUNCIONES DE DATOS
// ============================================
function pickFirstMatchingColumn(headers, candidates) {
  const set = new Set(headers.map(h => String(h).trim().toLowerCase()));
  return candidates.find(c => set.has(String(c).trim().toLowerCase())) || null;
}
function excelBufferToList(buffer, columnName) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

  return rows
    .map(r => String(r[columnName] ?? '').trim())
    .filter(Boolean);
}

function uniqSorted(arr) {
  return [...new Set(arr.map(x => String(x).trim()).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, 'es'));
}

async function readJsonArray(file, fallback = []) {
  try {
    if (!fs.existsSync(file)) return fallback;
    const raw = await fsp.readFile(file, 'utf8');
    const data = JSON.parse(raw.replace(/^\uFEFF/, '').trim());
    return Array.isArray(data) ? data : fallback;
  } catch {
    return fallback;
  }
}

async function asegurarLocationsFile() {
  try {
    if (!fs.existsSync(LOCATIONS_FILE)) {
      const initial = [
        { name: 'Beasain', coords: [43.04426527618791, -2.2100984760320834] },
        { name: 'Madrid', coords: [40.4168, -3.7038] },
        { name: 'Barcelona', coords: [41.3851, 2.1734] }
      ];
      await fsp.writeFile(LOCATIONS_FILE, JSON.stringify(initial, null, 2), 'utf8');
      console.log('✅ locations.json creado:', LOCATIONS_FILE);
    }
  } catch (e) {
    console.warn('⚠️ No se pudo crear locations.json:', e.message);
  }
}

async function inicializarDatos() {
  try {
    await fsp.access(DATA_FILE);
    console.log('✅ Archivo de datos existe:', DATA_FILE);
  } catch {
    console.log('📝 Creando archivo de datos inicial...');
    const datos = {
      viajes: [
        {
          id: 1,
          traveler: 'Ana García',
          destination: 'Madrid',
          startDate: '2026-01-19',
          endDate: '2026-01-21',
          client: 'Iberdrola',
          project: 'Instalación eólica offshore',
          color: '#c10230',
          coords: [40.4168, -3.7038],
          estado: 'Aprobado',
          creadoPor: 'admin',
          fechaCreacion: new Date().toISOString()
        }
      ],
      clientes: [
        { id: 1, nombre: 'Iberdrola' }
      ],
      config: {
        ultimoID: 1,
        version: '2.3.0',
        ultimaActualizacion: new Date().toISOString()
      }
    };
    await fsp.writeFile(DATA_FILE, JSON.stringify(datos, null, 2), 'utf8');
    console.log('✅ Datos inicializados con', datos.viajes.length, 'viajes de ejemplo');
  }

  // asegurar geocache
  try {
    if (!fs.existsSync(GEO_CACHE_FILE)) {
      await fsp.writeFile(GEO_CACHE_FILE, '{}', 'utf8');
    }
  } catch (e) {
    console.warn('⚠️ No se pudo inicializar geocache:', e.message);
  }
}

async function leerDatos() {
  const contenido = await fsp.readFile(DATA_FILE, 'utf8');
  return JSON.parse(contenido.replace(/^\uFEFF/, '').trim());
}

async function escribirDatos(datos) {
  await fileLock.acquire();
  try {
    await crearBackup();
    datos.config = datos.config || {};
    datos.config.ultimaActualizacion = new Date().toISOString();
    await fsp.writeFile(DATA_FILE, JSON.stringify(datos, null, 2), 'utf8');
  } finally {
    await fileLock.release(); 
  }
}

async function crearBackup() {
  try {
    if (!fs.existsSync(DATA_FILE)) return;

    const timestamp = new Date().toISOString().replace(/:/g, '-').split('.')[0];
    const backupPath = path.join(BACKUP_DIR, `viajes_backup_${timestamp}.json`);

    const contenido = await fsp.readFile(DATA_FILE, 'utf8');
    await fsp.writeFile(backupPath, contenido, 'utf8');

    // Mantener solo los últimos 100 backups
    const backups = await fsp.readdir(BACKUP_DIR);
    if (backups.length > 100) {
      const sorted = backups.sort();
      const toDelete = sorted.slice(0, backups.length - 100);
      for (const file of toDelete) {
        await fsp.unlink(path.join(BACKUP_DIR, file));
      }
    }
  } catch (e) {
    console.warn('⚠️ Error creando backup:', e.message);
  }
}

// Backup automático cada hora
setInterval(crearBackup, 60 * 60 * 1000);

// ============================================
// Util: limpiar nombres (clientes catálogo)
// ============================================

function limpiarNombreCliente(s) {
  if (!s) return '';
  let t = String(s).trim();

  // quitar BOM/zero width
  t = t.replace(/[\u200B-\u200D\uFEFF]/g, '');

  // si viene entrecomillado entero
  if (t.startsWith('"') && t.endsWith('"') && t.length >= 2) {
    t = t.slice(1, -1);
  }

  // csv escaping "" -> "
  t = t.replace(/""/g, '"');

  return t.trim();
}

// ============================================
// ENDPOINTS API REST
// ============================================

// Health
app.get('/api/health', (req, res) => {
  res.json({
    status: 'OK',
    timestamp: new Date().toISOString(),
    dataDir: DATA_DIR,
    esEjecutable
  });
});
// Catálogo clientes (para autocompletar)
app.get('/api/clientes', async (req, res) => {
  try {
    if (!fs.existsSync(CLIENTES_FILE)) {
      return res.status(404).json({ error: 'Archivo no encontrado', path: CLIENTES_FILE });
    }

    let raw = await fsp.readFile(CLIENTES_FILE, 'utf8');
    raw = raw.replace(/^\uFEFF/, '').trim();

    const clientes = JSON.parse(raw);
    const clean = Array.isArray(clientes)
      ? clientes
          .filter(c => c && c !== 'Account Name')
          .map(limpiarNombreCliente)
          .filter(c => c !== '')
      : [];

    res.json(clean);
  } catch (e) {
    console.error('❌ Error leyendo catálogo clientes:', e);
    res.status(500).json({ error: 'No se pudo leer el catálogo de clientes' });
  }
});

app.get('/api/ofertas', async (req, res) => {
  try {
    if (!fs.existsSync(OFERTAS_FILE)) return res.json([]);

    let raw = await fsp.readFile(OFERTAS_FILE, 'utf8');
    raw = raw.replace(/^\uFEFF/, '').trim();

    const ofertas = JSON.parse(raw);
    const clean = Array.isArray(ofertas)
      ? ofertas.map(x => String(x || '').trim()).filter(Boolean)
      : [];

    return res.json(clean);
  } catch (e) {
    console.error('❌ Error leyendo ofertas.json:', e);
    return res.status(500).json({ error: 'No se pudo leer el catálogo de ofertas' });
  }
});
app.get('/api/proyectos', async (req, res) => {
  const list = await readJsonArray(PROYECTOS_FILE, []);
  res.json(list);
});

// Viajes
app.get('/api/viajes', async (req, res) => {
  try {
    const datos = await leerDatos();
    res.json(datos.viajes || []);
  } catch (e) {
    console.error('Error GET /api/viajes:', e);
    res.status(500).json({ error: 'Error al obtener viajes' });
  }
});

app.get('/api/viajes/:id', async (req, res) => {
  try {
    const datos = await leerDatos();
    const v = (datos.viajes || []).find(x => x.id == req.params.id);
    if (!v) return res.status(404).json({ error: 'Viaje no encontrado' });
    res.json(v);
  } catch (e) {
    res.status(500).json({ error: 'Error al obtener viaje' });
  }
});
// GET /api/locations - catálogo de ubicaciones
app.get('/api/locations', async (req, res) => {
  try {
    if (!fs.existsSync(LOCATIONS_FILE)) return res.json([]);
    const raw = await fsp.readFile(LOCATIONS_FILE, 'utf8');
    const data = JSON.parse(raw.replace(/^\uFEFF/, '').trim());
    res.json(Array.isArray(data) ? data : []);
  } catch (e) {
    console.error('❌ Error leyendo locations:', e);
    res.status(500).json({ error: 'No se pudo leer locations' });
  }
});

// POST /api/locations - crear ubicación nueva
app.post('/api/locations', async (req, res) => {
  try {
    const { name, coords } = req.body;

    if (!name || !Array.isArray(coords) || coords.length !== 2) {
      return res.status(400).json({ error: 'Formato inválido. Se requiere {name, coords:[lat,lon]}' });
    }

    const lat = Number(coords[0]);
    const lon = Number(coords[1]);
    if (Number.isNaN(lat) || Number.isNaN(lon)) {
      return res.status(400).json({ error: 'Coordenadas inválidas' });
    }

    await fileLock.acquire();
    try {
      const raw = fs.existsSync(LOCATIONS_FILE) ? await fsp.readFile(LOCATIONS_FILE, 'utf8') : '[]';
      const list = JSON.parse(raw.replace(/^\uFEFF/, '').trim() || '[]');

      const normalized = name.trim().toLowerCase();
      const exists = Array.isArray(list) && list.some(l => (l.name || '').trim().toLowerCase() === normalized);
      if (exists) {
        return res.status(409).json({ error: 'La ubicación ya existe' });
      }

      const newLoc = { name: name.trim(), coords: [lat, lon] };
      const next = Array.isArray(list) ? [...list, newLoc] : [newLoc];

      await fsp.writeFile(LOCATIONS_FILE, JSON.stringify(next, null, 2), 'utf8');
      res.status(201).json(newLoc);
    } finally {
      await fileLock.release();
    }
  } catch (e) {
    console.error('❌ Error guardando location:', e);
    res.status(500).json({ error: 'No se pudo guardar location' });
  }
});
app.post('/api/viajes', async (req, res) => {
  try {
    const datos = await leerDatos();
    datos.config = datos.config || {};
    datos.viajes = datos.viajes || [];

    // Nuevo ID incremental
    datos.config.ultimoID = (datos.config.ultimoID || 0) + 1;
    const id = datos.config.ultimoID;

    const {
      traveler,
      destination,
      startDate,
      endDate,
      client,
      project,
      coords,
      userId,
      creadoPor,
      color
    } = req.body;

    // Validación mínima
    if (!traveler || !destination || !startDate || !endDate || !client || !project) {
      return res.status(400).json({ error: 'Faltan campos obligatorios' });
    }

    // Coordenadas: vienen del frontend (catálogo locations) o fallback
    let finalCoords = office;
    if (Array.isArray(coords) && coords.length === 2) {
      const lat = Number(coords[0]);
      const lon = Number(coords[1]);
      if (!Number.isNaN(lat) && !Number.isNaN(lon)) {
        finalCoords = [lat, lon];
      }
    }

    const nuevoViaje = {
      id,
      userId: userId || null,                 // <-- guardar userId
      creadoPor: (creadoPor || traveler).trim(), // <-- quien lo crea (usuario actual)
      traveler: traveler.trim(),              // <-- viajero (en tu caso = usuario)
      destination: destination.trim(),
      startDate,
      endDate,
      client: client.trim(),
      project: project.trim(),
      coords: finalCoords,
      color: color || '#636569',              // <-- usar color del usuario si viene
      estado: 'Aprobado',
      fechaCreacion: new Date().toISOString()
    };

    datos.viajes.push(nuevoViaje);
    await escribirDatos(datos);

    return res.status(201).json(nuevoViaje);
  } catch (error) {
    console.error('Error endpoint POST /api/viajes:', error);
    return res.status(500).json({ error: 'Error al crear viaje' });
  }
});

app.get('/api/users', async (req, res) => {
  try {
    if (!fs.existsSync(USERS_FILE)) return res.json([]);
    const raw = await fsp.readFile(USERS_FILE, 'utf8');
    const data = JSON.parse(raw.replace(/^\uFEFF/, '').trim());
    return res.json(Array.isArray(data) ? data : []);
  } catch (e) {
    console.error('❌ Error leyendo users.json:', e);
    return res.status(500).json({ error: 'No se pudo leer el catálogo de usuarios' });
  }
});
app.put('/api/viajes/:id', async (req, res) => {
  try {
    const datos = await leerDatos();
    datos.viajes = datos.viajes || [];

    const index = datos.viajes.findIndex(v => v.id == req.params.id);
    if (index === -1) {
      return res.status(404).json({ error: 'Viaje no encontrado' });
    }

    const viajeActual = datos.viajes[index];

    // Coords: si vienen válidas, se actualizan; si no, se mantienen
    let coordsActualizadas = viajeActual.coords;
    if (Array.isArray(req.body.coords) && req.body.coords.length === 2) {
      const lat = Number(req.body.coords[0]);
      const lon = Number(req.body.coords[1]);
      if (!Number.isNaN(lat) && !Number.isNaN(lon)) {
        coordsActualizadas = [lat, lon];
      }
    }

    // Color: solo actualizar si viene un string no vacío
    const colorActualizado =
      (typeof req.body.color === 'string' && req.body.color.trim())
        ? req.body.color.trim()
        : viajeActual.color;

    // userId: solo actualizar si viene (si no, mantener)
    const userIdActualizado =
      (typeof req.body.userId === 'string' && req.body.userId.trim()) || (req.body.userId === null)
        ? req.body.userId
        : viajeActual.userId;

        const viajeActualizado = {
          ...viajeActual,
          ...req.body,
        
          id: viajeActual.id,
          coords: coordsActualizadas,
        
          // PRESERVAR creador original
          traveler: viajeActual.traveler,
          userId: viajeActual.userId,
          creadoPor: viajeActual.creadoPor,
          color: viajeActual.color,
        
          // registrar editor
          modificadoPor: req.body.modificadoPor || viajeActual.modificadoPor,
          fechaModificacion: new Date().toISOString()
        };

    // Limpieza básica de strings si vienen
    if (typeof viajeActualizado.traveler === 'string') viajeActualizado.traveler = viajeActualizado.traveler.trim();
    if (typeof viajeActualizado.destination === 'string') viajeActualizado.destination = viajeActualizado.destination.trim();
    if (typeof viajeActualizado.client === 'string') viajeActualizado.client = viajeActualizado.client.trim();
    if (typeof viajeActualizado.project === 'string') viajeActualizado.project = viajeActualizado.project.trim();
    if (typeof viajeActualizado.creadoPor === 'string') viajeActualizado.creadoPor = viajeActualizado.creadoPor.trim();
    if (typeof viajeActualizado.modificadoPor === 'string') viajeActualizado.modificadoPor = viajeActualizado.modificadoPor.trim();

    datos.viajes[index] = viajeActualizado;
    await escribirDatos(datos);

    return res.json(viajeActualizado);
  } catch (error) {
    console.error('Error endpoint PUT /api/viajes/:id:', error);
    return res.status(500).json({ error: 'Error al actualizar viaje' });
  }
});

app.delete('/api/viajes/:id', async (req, res) => {
  try {
    const datos = await leerDatos();
    datos.viajes = datos.viajes || [];

    const idx = datos.viajes.findIndex(v => v.id == req.params.id);
    if (idx === -1) return res.status(404).json({ error: 'Viaje no encontrado' });

    const eliminado = datos.viajes.splice(idx, 1)[0];
    await escribirDatos(datos);

    res.json({ mensaje: 'Viaje eliminado', viaje: eliminado });
  } catch (e) {
    console.error('Error DELETE /api/viajes/:id:', e);
    res.status(500).json({ error: 'Error al eliminar viaje' });
  }
});

// Stats
app.get('/api/stats', async (req, res) => {
  try {
    const datos = await leerDatos();
    const viajes = datos.viajes || [];
    const stats = {
      totalViajes: viajes.length,
      viajeros: [...new Set(viajes.map(v => v.traveler))].length,
      clientes: [...new Set(viajes.map(v => v.client))].length
    };
    res.json(stats);
  } catch (e) {
    res.status(500).json({ error: 'Error al obtener estadísticas' });
  }
});

// Servir index.html
app.get('/', (req, res) => {
    const indexPath = path.join(__dirname, '../public/index.html');
  res.sendFile(indexPath);
});

// ============================================
// INICIAR SERVIDOR
// ============================================
const PORT = process.env.PORT || 3001;

async function start() {
  await inicializarDatos();
  await asegurarLocationsFile();
  app.listen(PORT, () => {
    console.log(`
----------------------------------------------
-   VIAJES CORPORATIVOS          -
-   Servidor: http://localhost:${PORT}      -
-   Datos: ${DATA_FILE}
-   Backups: ${BACKUP_DIR}
----------------------------------------------
    `);
  });
}

// Si lo lanza Electron con require(), también arranca
start().catch(err => {
  console.error('❌ No se pudo arrancar el servidor:', err);
  process.exit(1);
});

// ============================================
// MANEJO DE CIERRE GRACEFUL
// ============================================
process.on('SIGINT', async () => {
  console.log('\n🛑 Cerrando servidor...');
  await fileLock.release();
  await crearBackup();
  process.exit(0);
});

process.on('SIGTERM', async () => {
  console.log('\n🛑 Cerrando servidor...');
  await fileLock.release();
  await crearBackup();
  process.exit(0);
});
