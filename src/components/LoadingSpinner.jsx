export const LoadingSpinner = ({ mensaje = 'Cargando...' }) => (
  <div style={{
    display: 'flex', flexDirection: 'column',
    alignItems: 'center', justifyContent: 'center',
    minHeight: '60vh', gap: '16px',
  }}>
    <div style={{
      width: '40px', height: '40px',
      border: '3px solid #e0e4ea',
      borderTop: '3px solid #004895',
      borderRadius: '50%',
      animation: 'spin 0.75s linear infinite',
    }} />
    <p style={{ color: 'var(--text-2)', fontSize: '14px', fontWeight: 500 }}>{mensaje}</p>
    <style>{`@keyframes spin { to { transform: rotate(360deg) } }`}</style>
  </div>
)

export const FullscreenFeedback = ({ tipo, mensaje, onContinuar }) => (
  <div style={{
    position: 'fixed', inset: 0,
    backgroundColor: tipo === 'exito' ? '#004895' : '#c62828',
    display: 'flex', flexDirection: 'column',
    alignItems: 'center', justifyContent: 'center',
    zIndex: 200, gap: '20px', padding: '32px',
    animation: 'fadeIn 0.3s ease',
  }}>
    <div style={{
      width: '72px', height: '72px',
      backgroundColor: 'rgba(255,255,255,0.15)',
      borderRadius: '50%',
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      fontSize: '36px',
    }}>
      {tipo === 'exito' ? '✓' : '✗'}
    </div>
    <p style={{ color: 'white', fontSize: '20px', fontWeight: 700, textAlign: 'center' }}>{mensaje}</p>
    {onContinuar && (
      <button onClick={onContinuar} style={{
        marginTop: '8px', backgroundColor: '#F8A12F', color: 'white',
        border: 'none', borderRadius: '14px', padding: '16px 40px',
        fontSize: '17px', fontWeight: 700, minHeight: '56px',
      }}>
        Continuar
      </button>
    )}
    <style>{`@keyframes fadeIn { from { opacity:0 } to { opacity:1 } }`}</style>
  </div>
)
