import { tabletConfig } from '../config/tabletConfig'
import { useApp } from '../context/AppContext'

const BTN_ROL = {
  background: 'transparent',
  border: '1.5px solid rgba(255,255,255,0.4)',
  color: 'rgba(255,255,255,0.85)',
  borderRadius: '8px', padding: '5px 10px',
  fontSize: '11px', minHeight: '30px', cursor: 'pointer',
}

export const Header = ({ titulo, subtitulo, pendingCount = 0, onLogout, onCambiarRol, accion, color = '#004895' }) => {
  const { rol, seleccionarRol } = useApp()
  const mostrarRol = onCambiarRol || rol === 'bi'
  const handleRol = onCambiarRol || (() => seleccionarRol(null))

  return (
    <header style={{
      backgroundColor: color,
      color: 'white',
      padding: '10px 18px',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'space-between',
      position: 'sticky',
      top: 0,
      zIndex: 100,
      boxShadow: '0 2px 8px rgba(0,0,0,0.25)',
      minHeight: '58px',
    }}>
      <div style={{ minWidth: 0 }}>
        <p style={{ fontSize: '10px', opacity: 0.65, textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: '1px' }}>
          {tabletConfig.nombreMaquina}
        </p>
        <h1 style={{ fontSize: '16px', fontWeight: 700, lineHeight: 1.2, color: 'white' }}>
          {titulo}
        </h1>
        {subtitulo && (
          <p style={{ fontSize: '11px', opacity: 0.75, marginTop: '1px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
            {subtitulo}
          </p>
        )}
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0, marginLeft: '12px' }}>
        {pendingCount > 0 && (
          <div title={`${pendingCount} registro(s) pendientes`} style={{
            display: 'flex', alignItems: 'center', gap: '4px',
            backgroundColor: '#F8A12F',
            borderRadius: '20px', padding: '3px 10px',
            fontSize: '11px', fontWeight: 700,
          }}>
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
              <path d="M3 3l18 18M17.5 17.5A5 5 0 0 0 14 8h-1.26A8 8 0 1 0 5.07 15.29" />
            </svg>
            {pendingCount}
          </div>
        )}

        {accion}

        {mostrarRol && (
          <button onClick={handleRol} style={BTN_ROL}>
            ⇄ Rol
          </button>
        )}

        {onLogout && (
          <button onClick={() => {
            if (window.confirm('¿Seguro que quieres cerrar sesión?')) onLogout()
          }} style={{
            background: 'transparent',
            border: '1.5px solid rgba(255,255,255,0.4)',
            color: 'rgba(255,255,255,0.85)',
            borderRadius: '8px', padding: '5px 12px',
            fontSize: '12px', fontWeight: 600, minHeight: '32px', cursor: 'pointer',
          }}>
            Salir
          </button>
        )}
      </div>
    </header>
  )
}
