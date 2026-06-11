import { useState, useEffect, useCallback } from 'react'
import { getListItems, createListItem } from '../services/sharepoint'
import { encolarOperacion } from '../services/indexedDB'

// Maneja registros parciales de producción para el turno activo
export const useProduccionParcial = ({ getToken, registroId, codigoLote }) => {
  const [parciales, setParciales]       = useState([])
  const [baseAcumLote, setBaseAcumLote] = useState(0) // acumulado lote de registros anteriores
  const [cargando, setCargando]         = useState(false)

  // Totales del turno actual (solo parciales de este registro)
  const ultimoParcial = parciales[parciales.length - 1] || null
  const acumConfTurno = ultimoParcial?.Acum_Conf_Turno ?? 0
  const acumDefTurno  = ultimoParcial?.Acum_Def_Turno  ?? 0
  // Acumulado lote = base histórica + acumulado de este registro
  const acumConfLote  = baseAcumLote + (ultimoParcial?.Acum_Conf_Turno ?? 0)

  const cargar = useCallback(async () => {
    if (!registroId) return
    setCargando(true)
    try {
      const token = await getToken()
      if (!token) return

      // 1. Parciales del registro activo (para KPIs del turno)
      const items = await getListItems(token, 'Registro_Produccion_Parcial', {
        filter: `Registro_ID eq ${registroId}`,
        orderby: 'Timestamp asc',
      })
      setParciales(items)

      // 2. Último parcial de otro registro del mismo lote (para base acumulada del lote)
      if (codigoLote) {
        try {
          const loteParciales = await getListItems(token, 'Registro_Produccion_Parcial', {
            filter: `Codigo_Lote eq '${codigoLote}'`,
            orderby: 'Timestamp desc',
            top: 50,
          })
          // Buscar el último parcial de un registro DIFERENTE al actual
          const otroRegistro = loteParciales.find(p => p.Registro_ID !== registroId)
          if (otroRegistro) {
            setBaseAcumLote(otroRegistro.Acum_Conf_Lote || 0)
          }
        } catch { /* lote sin historial */ }
      }
    } catch {
      // Lista puede no tener datos aún
    } finally {
      setCargando(false)
    }
  }, [registroId, codigoLote])

  // eslint-disable-next-line react-hooks/set-state-in-effect
  useEffect(() => { cargar() }, [cargar])

  const agregarParcial = async ({ undConformes, undDefectuosas, operario, obs }) => {
    const token = await getToken()
    if (!token) return null

    // Acumulados del turno actual
    const newAcumConf = acumConfTurno + undConformes
    const newAcumDef  = acumDefTurno  + undDefectuosas
    // Acumulado lote = base histórica + acumulado turno + nuevo
    const newAcumLote = baseAcumLote + newAcumConf

    const payload = {
      Title:           `${registroId}-${new Date().toISOString().slice(11,16)}`,
      Registro_ID:     registroId,
      Codigo_Lote:     codigoLote || '',
      Timestamp:       new Date().toISOString(),
      Und_Conformes:   undConformes,
      Und_Defectuosas: undDefectuosas,
      Acum_Conf_Turno: newAcumConf,
      Acum_Def_Turno:  newAcumDef,
      Acum_Conf_Lote:  newAcumLote,
      Operario:        operario || '',
      Obs:             obs || '',
    }

    try {
      if (navigator.onLine) {
        const resultado = await createListItem(token, 'Registro_Produccion_Parcial', payload)
        const nuevo = { ...payload, ID: resultado.ID }
        setParciales(prev => [...prev, nuevo])
        return nuevo
      } else {
        throw new Error('offline')
      }
    } catch {
      // Fallback offline: encolar para sincronizar cuando haya conexión
      const localId = crypto.randomUUID()
      await encolarOperacion({
        tipo: 'create',
        listName: 'Registro_Produccion_Parcial',
        data: payload,
        localId,
      })
      const nuevo = { ...payload, ID: localId, _offline: true }
      setParciales(prev => [...prev, nuevo])
      return nuevo
    }
  }

  const ultimoTimestamp = ultimoParcial?.Timestamp || null

  return {
    parciales,
    acumConfTurno,
    acumDefTurno,
    acumConfLote,
    ultimoTimestamp,
    cargando,
    cargar,
    agregarParcial,
  }
}
