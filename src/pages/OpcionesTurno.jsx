import { useMsal } from '../hooks/useMsal'
import { useApp } from '../context/AppContext'
import { updateListItem } from '../services/sharepoint'
import { tabletConfig } from '../config/tabletConfig'

const Boton = ({ onClick, color, titulo, descripcion }) => (
  <button onClick={onClick} style={{
    backgroundColor: color, color: 'white', border: 'none',
    borderRadius: '16px', padding: '18px 20px', fontSize: '16px',
    fontWeight: 700, minHeight: '72px', cursor: 'pointer',
    lineHeight: 1.3, textAlign: 'left', width: '100%',
    boxShadow: '0 3px 10px rgba(0,0,0,0.2)',
  }}>
    {titulo}
    {descripcion && (
      <div style={{ fontSize: '12px', opacity: 0.85, fontWeight: 400, marginTop: '4px' }}>
        {descripcion}
      </div>
    )}
  </button>
)

export default function OpcionesTurno() {
  const { getToken, logout } = useMsal()
  const { turnoActivo, limpiarTurno, setPantalla, setModoRelevo, actualizarTurnoLocal } = useApp()

  const marcarTransferido = async () => {
    try {
      const token = await getToken()
      if (token && turnoActivo?.spId && navigator.onLine) {
        await updateListItem(token, 'Registro_Produccion', turnoActivo.spId, { Estado: 'transferido' })
      }
      await actualizarTurnoLocal({ Estado: 'transferido' })
    } catch { /* offline — PCP puede corregir */ }
  }

  return (
    <div style={{
      position: 'fixed', inset: 0, backgroundColor: '#004895',
      display: 'flex', flexDirection: 'column',
      alignItems: 'center', justifyContent: 'center',
      padding: '32px', gap: '16px',
    }}>
      {/* Encabezado */}
      <div style={{ textAlign: 'center', marginBottom: '8px' }}>
        <div style={{ fontSize: '48px', marginBottom: '8px' }}>✓</div>
        <h2 style={{ color: 'white', fontSize: '22px', fontWeight: 800, margin: 0 }}>
          Turno finalizado
        </h2>
        <p style={{ color: 'rgba(255,255,255,0.7)', fontSize: '14px', marginTop: '6px' }}>
          {tabletConfig.nombreMaquina} · {turnoActivo?.Producto} · {turnoActivo?.Turno}
        </p>
      </div>

      <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', width: '100%', maxWidth: '360px' }}>

        {/* Opción 1 — Relevo: mismo producto, operario entrante (Caso A — Manual Jefe Op. N2) */}
        <Boton
          color="#37BEEC"
          titulo="🔄 Continuar este producto"
          descripcion="El operario entrante retoma la misma producción — se pre-carga producto y color"
          onClick={async () => {
            await marcarTransferido()  // Estado → transferido (Caso A del manual)
            setModoRelevo(true)
            setPantalla('cambio-producto')
          }}
        />

        {/* Opción 2 — Nuevo producto */}
        <Boton
          color="#1565c0"
          titulo="📦 Cambiar de producto"
          descripcion="La máquina continúa pero con un producto diferente"
          onClick={() => {
            setModoRelevo(false)
            setPantalla('cambio-producto')
          }}
        />

        {/* Opción 3 — Cierre definitivo */}
        <Boton
          color="#F8A12F"
          titulo="🆕 Producción finalizada"
          descripcion="Nuevo molde y nueva MP al retomar — vuelves a la pantalla principal"
          onClick={limpiarTurno}
        />

      </div>

      {/* Salir */}
      <button onClick={logout} style={{
        marginTop: '8px', backgroundColor: 'transparent',
        color: 'rgba(255,255,255,0.5)', border: '1px solid rgba(255,255,255,0.3)',
        borderRadius: '8px', padding: '8px 20px', fontSize: '13px', cursor: 'pointer',
      }}>
        Cerrar sesión
      </button>
    </div>
  )
}
