import { useEffect } from 'react'
import { createPortal } from 'react-dom'

const COLORES = {
  error: { bg: '#c62828', borde: '#b71c1c' },
  ok:    { bg: '#2e7d32', borde: '#1b5e20' },
  warn:  { bg: '#e65100', borde: '#bf360c' },
  info:  { bg: '#004895', borde: '#002d5a' },
}

// Uso: const [toast, setToast] = useState(null)
//      setToast({ mensaje: 'Texto', tipo: 'error' | 'ok' | 'warn' | 'info' })
//      <Toast toast={toast} onClose={() => setToast(null)} />
export const Toast = ({ toast, onClose }) => {
  useEffect(() => {
    if (!toast) return
    const t = setTimeout(onClose, toast.tipo === 'error' ? 5000 : 3000)
    return () => clearTimeout(t)
  }, [toast])

  if (!toast) return null
  const c = COLORES[toast.tipo] || COLORES.error

  return createPortal(
    <div style={{
      position: 'fixed', bottom: '28px', left: '50%', transform: 'translateX(-50%)',
      zIndex: 999999, maxWidth: '400px', width: 'calc(100% - 32px)',
      backgroundColor: c.bg, border: `1.5px solid ${c.borde}`,
      borderRadius: '12px', padding: '14px 16px',
      color: 'white', fontSize: '14px', fontWeight: 600,
      boxShadow: '0 6px 24px rgba(0,0,0,0.35)',
      display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '10px',
    }}>
      <span style={{ lineHeight: 1.4 }}>{toast.mensaje}</span>
      <button onClick={onClose} style={{
        background: 'none', border: 'none', color: 'rgba(255,255,255,0.8)',
        fontSize: '20px', cursor: 'pointer', padding: 0, lineHeight: 1, flexShrink: 0,
      }}>×</button>
    </div>,
    document.body
  )
}

export const mensajeRed = (err) =>
  (err?.name === 'TypeError' || err?.message === 'Failed to fetch')
    ? 'Sin conexión. Verifica tu red e intenta de nuevo.'
    : ('Error: ' + (err?.message || 'desconocido'))
