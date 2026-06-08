import { useState, useEffect } from 'react'
import { MAQUINAS_DISPONIBLES, guardarTabletConfig, guardarTabletConfigTemporal, getTabletToken } from '../config/tabletConfig'
import { getListItems, updateListItem } from '../services/sharepoint'
import { useMsal } from '../hooks/useMsal'

export default function ConfigurarTablet({ onConfigurado, onTemporal }) {
  const { getToken } = useMsal()
  const [seleccion, setSeleccion]         = useState(null)
  const [asignaciones, setAsignaciones]   = useState({}) // { codigoMaquina: { spId, token } }
  const [guardando, setGuardando]         = useState(false)
  const [cargando, setCargando]           = useState(true)
  const [error, setError]                 = useState(null)

  const miToken = getTabletToken()

  // Cargar asignaciones actuales desde Maestro_Maquinas
  useEffect(() => {
    const cargar = async () => {
      try {
        const token = await getToken()
        if (!token) return
        const maquinas = await getListItems(token, 'Maestro_Maquinas')
        const map = {}
        maquinas.forEach(m => {
          const codigo = m.Codigo || m.Title || ''
          if (codigo) map[codigo] = { spId: m.ID, tabletToken: m.TabletToken || null, tabletFecha: m.TabletFecha || null }
        })
        setAsignaciones(map)
      } catch {
        // Sin conexión — permitir configurar igual (verificará al subir)
      } finally {
        setCargando(false)
      }
    }
    cargar()
  }, [])

  const esMia = (codigo) => asignaciones[codigo]?.tabletToken === miToken
  const estaAsignada = (codigo) => !!asignaciones[codigo]?.tabletToken && !esMia(codigo)

  const confirmar = async () => {
    if (!seleccion) return
    setGuardando(true)
    setError(null)
    try {
      const token = await getToken()
      if (token && asignaciones[seleccion.codigoMaquina]?.spId) {
        await updateListItem(token, 'Maestro_Maquinas', asignaciones[seleccion.codigoMaquina].spId, {
          TabletToken: miToken,
          TabletFecha: new Date().toISOString(),
        })
      }
      guardarTabletConfig(seleccion.codigoMaquina, seleccion.nombreMaquina)
      onConfigurado()
    } catch (err) {
      // Sin conexión: guardar igual localmente, SP se actualiza después
      guardarTabletConfig(seleccion.codigoMaquina, seleccion.nombreMaquina)
      onConfigurado()
    } finally {
      setGuardando(false)
    }
  }

  return (
    <div style={{ minHeight: '100vh', backgroundColor: '#004895', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: '24px' }}>
      <div style={{ width: '100%', maxWidth: '400px' }}>
        <div style={{ textAlign: 'center', marginBottom: '28px', color: 'white' }}>
          <p style={{ fontSize: '12px', opacity: 0.6, textTransform: 'uppercase', letterSpacing: '0.1em', marginBottom: '8px' }}>
            JJ PLASTIC SAC
          </p>
          <h1 style={{ fontSize: '22px', fontWeight: 800, margin: 0 }}>Configurar tablet</h1>
          <p style={{ fontSize: '13px', opacity: 0.7, marginTop: '8px' }}>
            Selecciona la máquina asignada a esta tablet.
          </p>
        </div>

        {cargando ? (
          <p style={{ color: 'rgba(255,255,255,0.6)', textAlign: 'center' }}>Verificando asignaciones...</p>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginBottom: '20px' }}>
            {MAQUINAS_DISPONIBLES.map(m => {
              const asignada = estaAsignada(m.codigoMaquina)
              const mia = esMia(m.codigoMaquina)
              const sel = seleccion?.codigoMaquina === m.codigoMaquina
              return (
                <button key={m.codigoMaquina}
                  onClick={() => !asignada && setSeleccion(m)}
                  disabled={asignada}
                  style={{
                    padding: '14px 18px', borderRadius: '12px', border: '2px solid',
                    borderColor: sel ? 'white' : mia ? '#81C784' : asignada ? 'rgba(255,255,255,0.1)' : 'rgba(255,255,255,0.25)',
                    backgroundColor: sel ? 'white' : mia ? 'rgba(129,199,132,0.15)' : asignada ? 'rgba(0,0,0,0.2)' : 'rgba(255,255,255,0.08)',
                    color: sel ? '#004895' : asignada ? 'rgba(255,255,255,0.3)' : 'white',
                    textAlign: 'left', cursor: asignada ? 'not-allowed' : 'pointer',
                    display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                    opacity: asignada ? 0.5 : 1,
                  }}>
                  <span style={{ fontWeight: 700, fontSize: '14px' }}>{m.nombreMaquina}</span>
                  <span style={{ fontSize: '11px', opacity: 0.7, fontFamily: 'monospace' }}>
                    {mia ? '✓ Este tablet' : asignada ? '🔒 Asignada' : m.codigoMaquina}
                  </span>
                </button>
              )
            })}
          </div>
        )}

        {error && <p style={{ color: '#ffcc80', fontSize: '13px', textAlign: 'center', marginBottom: '12px' }}>{error}</p>}

        {/* Botón permanente */}
        <button onClick={confirmar} disabled={!seleccion || guardando} style={{
          width: '100%', padding: '14px', borderRadius: '12px', border: 'none',
          backgroundColor: seleccion ? '#F8A12F' : 'rgba(255,255,255,0.15)',
          color: 'white', fontSize: '15px', fontWeight: 800,
          cursor: seleccion && !guardando ? 'pointer' : 'not-allowed', minHeight: '52px',
        }}>
          {guardando ? '⏳ Guardando...' : seleccion ? `✓ Asignar permanentemente` : 'Selecciona una máquina'}
        </button>

        {/* Botón temporal — guarda en sessionStorage y recarga (no toca SP ni localStorage) */}
        {onTemporal && seleccion && (
          <button onClick={() => {
            guardarTabletConfigTemporal(seleccion.codigoMaquina, seleccion.nombreMaquina)
            window.location.reload()
          }} style={{
              width: '100%', padding: '12px', borderRadius: '12px',
              border: '1.5px solid rgba(255,255,255,0.35)',
              backgroundColor: 'transparent', color: 'rgba(255,255,255,0.85)',
              fontSize: '14px', fontWeight: 600, cursor: 'pointer', marginTop: '6px',
            }}>
            👁 Solo esta sesión (no reclama la máquina)
          </button>
        )}

        <p style={{ color: 'rgba(255,255,255,0.4)', fontSize: '11px', textAlign: 'center', marginTop: '14px' }}>
          Las máquinas en 🔒 ya están asignadas a otro tablet.<br/>
          Para liberarlas, usa el panel BI → Gestión de tablets.
        </p>
      </div>
    </div>
  )
}
