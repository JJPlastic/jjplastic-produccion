import { createContext, useContext, useState, useEffect, useCallback } from 'react'
import { guardarTurnoActivo, obtenerTurnoActivo, contarPendientes } from '../services/indexedDB'

const AppContext = createContext(null)

export const AppProvider = ({ children }) => {
  const [pantalla, setPantalla] = useState('bienvenida')
  const [turnoActivo, setTurnoActivoState] = useState(null)
  const [pendingCount, setPendingCount] = useState(0)
  const [cargandoInicial, setCargandoInicial] = useState(true)
  // rol: rol de navegación actual (puede cambiar al usar ⇄ Rol)
  // rolCuenta: rol real de la cuenta en Maestro_Operarios (fijo por sesión)
  const [rol, setRol]               = useState(null)
  const [rolCuenta, setRolCuenta]   = useState(null)
  const [modoRelevo, setModoRelevo] = useState(false)
  const [ofPreseleccionada, setOfPreseleccionada] = useState(null)
  const [productoPreseleccionado, setProductoPreseleccionado] = useState(null)

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
      rol, seleccionarRol, rolCuenta, setRolCuenta,
      modoRelevo, setModoRelevo,
      ofPreseleccionada, setOfPreseleccionada,
      productoPreseleccionado, setProductoPreseleccionado,
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
