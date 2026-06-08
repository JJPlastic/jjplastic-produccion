import { createContext, useContext, useState, useEffect, useCallback } from 'react'
import { guardarTurnoActivo, obtenerTurnoActivo, contarPendientes } from '../services/indexedDB'

const AppContext = createContext(null)

export const AppProvider = ({ children }) => {
  const [pantalla, setPantalla] = useState('bienvenida')
  const [turnoActivo, setTurnoActivoState] = useState(null)
  const [pendingCount, setPendingCount] = useState(0)
  const [cargandoInicial, setCargandoInicial] = useState(true)
  // 'operario' | 'pcp' | 'bi' | 'jefeoperaciones' | 'gerencia' | null
  // null = pendiente de detección desde Maestro_Operarios
  const [rol, setRol]               = useState(null)
  const [modoRelevo, setModoRelevo] = useState(false)
  const [ofPreseleccionada, setOfPreseleccionada] = useState(null) // OF elegida en Bienvenida → pasa a InicioTurno

  useEffect(() => {
    const init = async () => {
      try {
        const turno = await obtenerTurnoActivo()
        if (turno?.spId || turno?.localId) {
          setTurnoActivoState(turno)
          setPantalla('turno-activo')
        }
        const count = await contarPendientes()
        setPendingCount(count)
      } finally {
        setCargandoInicial(false)
      }
    }
    init()
  }, [])

  const refreshPendingCount = useCallback(async () => {
    const count = await contarPendientes()
    setPendingCount(count)
  }, [])

  const setTurnoActivo = useCallback(async (turno) => {
    await guardarTurnoActivo(turno)
    setTurnoActivoState(turno)
    setPantalla(turno ? 'turno-activo' : 'bienvenida')
  }, [])

  const irACambioProducto = useCallback(() => {
    setPantalla('cambio-producto')
  }, [])

  const actualizarTurnoLocal = useCallback(async (cambios) => {
    setTurnoActivoState((prev) => {
      const actualizado = { ...prev, ...cambios }
      guardarTurnoActivo(actualizado)
      return actualizado
    })
  }, [])

  const limpiarTurno = useCallback(async () => {
    await guardarTurnoActivo(null)
    setTurnoActivoState(null)
    setPantalla('bienvenida')
  }, [])

  const seleccionarRol = useCallback((nuevoRol) => {
    setRol(nuevoRol)
  }, [])

  return (
    <AppContext.Provider value={{
      pantalla, setPantalla,
      turnoActivo, setTurnoActivo, actualizarTurnoLocal, limpiarTurno, irACambioProducto,
      pendingCount, refreshPendingCount,
      cargandoInicial,
      rol, seleccionarRol,
      modoRelevo, setModoRelevo,
      ofPreseleccionada, setOfPreseleccionada,
    }}>
      {children}
    </AppContext.Provider>
  )
}

// eslint-disable-next-line react-refresh/only-export-components
export const useApp = () => {
  const ctx = useContext(AppContext)
  if (!ctx) throw new Error('useApp debe usarse dentro de AppProvider')
  return ctx
}
