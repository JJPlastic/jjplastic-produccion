import { openDB } from 'idb'

const DB_NAME = 'jjplastic-db'
const DB_VERSION = 2

const getDB = () =>
  openDB(DB_NAME, DB_VERSION, {
    upgrade(db, oldVersion) {
      if (oldVersion < 1) {
        db.createObjectStore('registros-pendientes', { keyPath: 'localId' })
      }
      if (oldVersion < 2) {
        // Caché de datos maestros (operarios, productos, motivos)
        db.createObjectStore('maestros-cache', { keyPath: 'key' })
        // Turno activo — un solo registro con key fija
        db.createObjectStore('turno-activo', { keyPath: 'key' })
      }
    },
  })

// ─── Sync queue ──────────────────────────────────────────────────────────────

export const encolarOperacion = async (operacion) => {
  const db = await getDB()
  await db.put('registros-pendientes', {
    localId: crypto.randomUUID(),
    status: 'pending',
    timestamp: Date.now(),
    intentos: 0,
    ...operacion,
  })
}

export const obtenerPendientes = async () => {
  const db = await getDB()
  const todos = await db.getAll('registros-pendientes')
  return todos.filter((r) => r.status === 'pending').sort((a, b) => a.timestamp - b.timestamp)
}

export const marcarSincronizado = async (localId) => {
  const db = await getDB()
  await db.delete('registros-pendientes', localId)
}

export const marcarError = async (localId) => {
  const db = await getDB()
  const item = await db.get('registros-pendientes', localId)
  if (item) await db.put('registros-pendientes', { ...item, status: 'error', intentos: (item.intentos || 0) + 1 })
}

export const contarPendientes = async () => {
  const pendientes = await obtenerPendientes()
  return pendientes.length
}

// ─── Turno activo ────────────────────────────────────────────────────────────

export const guardarTurnoActivo = async (turno) => {
  const db = await getDB()
  if (turno === null) {
    await db.delete('turno-activo', 'current')
  } else {
    await db.put('turno-activo', { key: 'current', ...turno })
  }
}

export const obtenerTurnoActivo = async () => {
  const db = await getDB()
  const row = await db.get('turno-activo', 'current')
  if (!row) return null
  const { key: _key, ...turno } = row
  return turno
}

// ─── Caché de maestros ───────────────────────────────────────────────────────

const CACHE_TTL = 1000 * 60 * 60 * 8 // 8 horas

export const guardarMaestro = async (nombre, datos) => {
  const db = await getDB()
  await db.put('maestros-cache', { key: nombre, datos, timestamp: Date.now() })
}

export const obtenerMaestro = async (nombre) => {
  const db = await getDB()
  const row = await db.get('maestros-cache', nombre)
  if (!row) return null
  if (Date.now() - row.timestamp > CACHE_TTL) return null // caché vencida
  return row.datos
}
