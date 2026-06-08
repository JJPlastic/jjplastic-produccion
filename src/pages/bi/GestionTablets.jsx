import { useState, useEffect } from 'react'
import { useMsal } from '../../hooks/useMsal'
import { getListItems, updateListItem } from '../../services/sharepoint'
import { Header } from '../../components/Header'

export default function GestionTablets({ onVolver, onLogout }) {
  const { getToken } = useMsal()
  const [maquinas, setMaquinas]   = useState([])
  const [cargando, setCargando]   = useState(true)
  const [liberando, setLiberando] = useState(null)
  const [mensaje, setMensaje]     = useState(null)

  const cargar = async () => {
    setCargando(true)
    try {
      const token = await getToken()
      if (!token) return
      const items = await getListItems(token, 'Maestro_Maquinas')
      setMaquinas(items)
    } catch { } finally { setCargando(false) }
  }

  useEffect(() => { cargar() }, [])

  const liberar = async (m) => {
    if (!window.confirm(`¿Liberar "${m.Nombre || m.Title}" para que otro tablet pueda usarla?`)) return
    setLiberando(m.ID)
    setMensaje(null)
    try {
      const token = await getToken()
      await updateListItem(token, 'Maestro_Maquinas', m.ID, { TabletToken: null, TabletFecha: null })
      setMensaje(`✓ ${m.Nombre || m.Title} liberada correctamente.`)
      cargar()
    } catch (err) {
      setMensaje(`Error: ${err.message}`)
    } finally { setLiberando(null) }
  }

  const asignadas = maquinas.filter(m => m.TabletToken)
  const libres    = maquinas.filter(m => !m.TabletToken)

  return (
    <div style={{ backgroundColor: '#f0f2f5', minHeight: '100vh' }}>
      <Header titulo="Gestión de tablets" pendingCount={0} onLogout={onLogout} color="#004895" />
      <div style={{ padding: '16px', maxWidth: '520px', margin: '0 auto', display: 'flex', flexDirection: 'column', gap: '14px' }}>

        <button onClick={onVolver} style={{ background: 'none', border: 'none', color: '#004895', fontSize: '14px', cursor: 'pointer', textAlign: 'left', padding: 0, fontWeight: 600 }}>
          ← Volver al panel PCP
        </button>

        {mensaje && (
          <div style={{ backgroundColor: '#e8f5e9', border: '1px solid #4caf50', borderRadius: '8px', padding: '10px 14px', fontSize: '13px', color: '#1b5e20', fontWeight: 600 }}>
            {mensaje}
          </div>
        )}

        {cargando ? <p style={{ color: '#888', textAlign: 'center', padding: '32px' }}>Cargando...</p> : (
          <>
            <div>
              <p style={{ fontSize: '11px', fontWeight: 700, color: '#555', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: '8px' }}>
                🔒 Asignadas a un tablet ({asignadas.length})
              </p>
              {asignadas.length === 0 && <p style={{ color: '#aaa', fontSize: '13px' }}>Ninguna máquina asignada.</p>}
              {asignadas.map(m => (
                <div key={m.ID} style={{ backgroundColor: 'white', borderRadius: '12px', padding: '12px 16px', border: '1.5px solid #e0e0e0', marginBottom: '8px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    <p style={{ fontWeight: 700, fontSize: '14px', color: '#1a1a1a', margin: 0 }}>{m.Nombre || m.Title}</p>
                    <p style={{ fontSize: '11px', color: '#888', margin: '2px 0 0' }}>
                      {m.Codigo || ''} · Asignada {m.TabletFecha ? new Date(m.TabletFecha).toLocaleDateString('es-PE') : '—'}
                    </p>
                  </div>
                  <button onClick={() => liberar(m)} disabled={liberando === m.ID} style={{
                    backgroundColor: '#ffebee', color: '#c62828', border: '1px solid #ef9a9a',
                    borderRadius: '8px', padding: '8px 14px', fontSize: '12px', fontWeight: 700,
                    cursor: liberando === m.ID ? 'not-allowed' : 'pointer', flexShrink: 0,
                  }}>
                    {liberando === m.ID ? '⏳' : '🔓 Liberar'}
                  </button>
                </div>
              ))}
            </div>

            <div>
              <p style={{ fontSize: '11px', fontWeight: 700, color: '#555', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: '8px' }}>
                ✓ Disponibles para asignar ({libres.length})
              </p>
              {libres.map(m => (
                <div key={m.ID} style={{ backgroundColor: '#f9fdf9', borderRadius: '10px', padding: '10px 16px', border: '1.5px dashed #c8e6c9', marginBottom: '6px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <p style={{ fontWeight: 600, fontSize: '13px', color: '#555', margin: 0 }}>{m.Nombre || m.Title}</p>
                  <span style={{ fontSize: '11px', color: '#2e7d32', fontWeight: 600 }}>✓ Libre</span>
                </div>
              ))}
            </div>
          </>
        )}
      </div>
    </div>
  )
}
