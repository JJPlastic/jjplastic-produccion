import { useState, useEffect } from 'react'
import { getListItems } from '../services/sharepoint'
import { guardarMaestro, obtenerMaestro } from '../services/indexedDB'

// Carga datos maestros con caché offline (8h TTL en IndexedDB)
export const useMaestros = (getToken) => {
  const [operarios, setOperarios]           = useState([])
  const [productos, setProductos]           = useState([])   // TipoProducto = PT
  const [materiasPrimasRaw, setMPRaw]       = useState([])   // TipoProducto = MP
  const [motivos, setMotivos]               = useState([])
  const [colores, setColores]               = useState([])
  const [cargando, setCargando]             = useState(true)
  const [error, setError]                   = useState(null)

  useEffect(() => {
    let cancelled = false
    const cargar = async () => {
      setCargando(true)
      setError(null)
      try {
        // Intentar desde caché primero (funciona offline)
        const [cachedOps, cachedProds, cachedMotivos, cachedColores] = await Promise.all([
          obtenerMaestro('operarios'),
          obtenerMaestro('productos'),
          obtenerMaestro('motivos'),
          obtenerMaestro('colores'),
        ])

        if (cachedOps) setOperarios(cachedOps)
        if (cachedProds) setProductos(cachedProds)
        if (cachedMotivos) setMotivos(cachedMotivos)
        if (cachedColores) setColores(cachedColores)
        const cachedMP = await obtenerMaestro('materias_primas')
        if (cachedMP) setMPRaw(cachedMP)

        // Si hay conexión, refrescar desde SharePoint
        // Promise.allSettled: si una lista falla, las demás siguen cargando
        if (navigator.onLine) {
          const token = await getToken()
          if (!token) return
          // Sin filtros OData — filtramos en cliente para evitar errores de compatibilidad
          const [rOps, rProds, rMots, rCols] = await Promise.allSettled([
            getListItems(token, 'Maestro_Operarios'),
            getListItems(token, 'Maestro_Productos', { top: 500 }),
            getListItems(token, 'Maestro_Motivos_Parada'),
            getListItems(token, 'Maestro_Colores'),
          ])
          if (cancelled) return

          const esActivo = (item) =>
            item.Activo === true || item.Activo === 1 || item.Activo === undefined

          const logFallo = (nombre, reason) => {
            if (reason?.name === 'AbortError') return
            if (reason?.name === 'TypeError') return // Failed to fetch / offline — caché activo
            console.warn(`${nombre} no disponible — usando caché:`, reason?.message || reason)
          }

          if (rOps.status === 'fulfilled') {
            const v = rOps.value.filter(esActivo)
            setOperarios(v)
            await guardarMaestro('operarios', v)
          } else { logFallo('Maestro_Operarios', rOps.reason) }

          if (rProds.status === 'fulfilled') {
            const todos = rProds.value
            const pt = todos.filter(p => esActivo(p) && (p.TipoProducto || '').toUpperCase() === 'PT')
            const mp = todos.filter(p => esActivo(p) && ['MP','MC'].includes((p.TipoProducto || '').toUpperCase()))
            setProductos(pt)
            setMPRaw(mp)
            await guardarMaestro('productos', pt)
            await guardarMaestro('materias_primas', mp)
          } else { logFallo('Maestro_Productos', rProds.reason) }

          if (rMots.status === 'fulfilled') {
            const v = rMots.value.filter(esActivo)
            setMotivos(v)
            await guardarMaestro('motivos', v)
          } else { logFallo('Maestro_Motivos_Parada', rMots.reason) }

          if (rCols.status === 'fulfilled') {
            const v = rCols.value.filter(esActivo)
            setColores(v)
            await guardarMaestro('colores', v)
          } else {
            // Maestro_Colores es opcional — colores se ingresan como texto libre si no existe
            if (rCols.reason?.name !== 'AbortError') {
              // Solo advertir si es error real (no 404 = lista no existe)
            }
          }
        }
      } catch (err) {
        if (!cancelled) setError(err.message)
      } finally {
        if (!cancelled) setCargando(false)
      }
    }
    cargar()
    return () => { cancelled = true }
  }, [])

  return { operarios, productos, materiasPrimas: materiasPrimasRaw, motivos, colores, cargando, error }
}
