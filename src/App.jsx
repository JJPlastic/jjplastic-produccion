import { lazy, Suspense, useState, useEffect } from 'react'
import { AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react'
import { useMsal } from './hooks/useMsal'
import { AppProvider, useApp } from './context/AppContext'
import { useSyncQueue } from './hooks/useSyncQueue'
import { LoadingSpinner } from './components/LoadingSpinner'
import { getListItems } from './services/sharepoint'
import { tabletConfigGuardada } from './config/tabletConfig'
import ConfigurarTablet from './pages/ConfigurarTablet'

const InicioTurno   = lazy(() => import('./pages/InicioTurno'))
const TurnoActivo   = lazy(() => import('./pages/TurnoActivo'))
const ParadaMaquina = lazy(() => import('./pages/ParadaMaquina'))
const CierreTurno   = lazy(() => import('./pages/CierreTurno'))
const ValidacionPCP  = lazy(() => import('./pages/pcp/ValidacionPCP'))
const KardexMP       = lazy(() => import('./pages/pcp/KardexMP'))
const CambioProducto      = lazy(() => import('./pages/CambioProducto'))
const OpcionesTurno       = lazy(() => import('./pages/OpcionesTurno'))
const BienvenidaOperario  = lazy(() => import('./pages/BienvenidaOperario'))
const DashboardGerencia   = lazy(() => import('./pages/gerencia/DashboardGerencia'))

const Login = () => {
  const { login } = useMsal()
  return (
    <div style={{
      display: 'flex', flexDirection: 'column',
      alignItems: 'center', justifyContent: 'center',
      minHeight: '100vh',
      background: 'linear-gradient(160deg, #003570 0%, #004895 60%, #005db0 100%)',
      gap: '40px', padding: '32px',
    }}>
      <div style={{ textAlign: 'center', animation: 'fadeIn 0.4s ease' }}>
        <div style={{
          width: '88px', height: '88px',
          background: 'linear-gradient(135deg, #F8A12F, #e6901e)',
          borderRadius: '24px',
          margin: '0 auto 24px',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          fontSize: '40px',
          boxShadow: '0 8px 24px rgba(248,161,47,0.4)',
        }}>🏭</div>
        <h1 style={{ color: 'white', fontSize: '26px', fontWeight: 800, marginBottom: '6px', letterSpacing: '-0.5px' }}>
          JJ PLASTIC S.A.C.
        </h1>
        <p style={{ color: 'rgba(255,255,255,0.6)', fontSize: '14px' }}>
          Sistema de registro de producción
        </p>
      </div>

      <div style={{ width: '100%', maxWidth: '300px', display: 'flex', flexDirection: 'column', gap: '12px' }}>
        <button onClick={login} style={{
          background: 'linear-gradient(135deg, #F8A12F, #e6901e)',
          color: 'white', border: 'none',
          borderRadius: '14px', padding: '18px',
          fontSize: '16px', fontWeight: 700,
          minHeight: '58px', width: '100%',
          boxShadow: '0 6px 20px rgba(248,161,47,0.4)',
          letterSpacing: '0.01em',
        }}>
          Iniciar sesión con Microsoft
        </button>
        <p style={{ color: 'rgba(255,255,255,0.35)', fontSize: '11px', textAlign: 'center' }}>
          Cuenta corporativa JJ Plastic
        </p>
      </div>
      <style>{`@keyframes fadeIn { from { opacity:0; transform:translateY(12px) } to { opacity:1; transform:translateY(0) } }`}</style>
    </div>
  )
}

// Pantalla de selección de rol
const SelectorRol = () => {
  const { seleccionarRol } = useApp()
  const { usuario, logout } = useMsal()
  return (
    <div style={{
      minHeight: '100vh',
      background: 'linear-gradient(160deg, #003570 0%, #004895 60%, #005db0 100%)',
      display: 'flex', flexDirection: 'column',
      alignItems: 'center', justifyContent: 'center',
      padding: '32px',
    }}>
      {/* Avatar / nombre */}
      <div style={{ textAlign: 'center', color: 'white', marginBottom: '32px' }}>
        <div style={{
          width: '56px', height: '56px',
          backgroundColor: 'rgba(255,255,255,0.15)',
          borderRadius: '50%',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          fontSize: '22px', margin: '0 auto 12px',
          border: '2px solid rgba(255,255,255,0.25)',
        }}>
          {usuario?.nombre?.charAt(0)?.toUpperCase() || '👤'}
        </div>
        <p style={{ fontSize: '15px', fontWeight: 600, opacity: 0.9 }}>{usuario?.nombre}</p>
        <p style={{ fontSize: '12px', opacity: 0.5, marginTop: '2px' }}>{usuario?.email}</p>
      </div>

      <h2 style={{ color: 'white', fontSize: '18px', fontWeight: 700, marginBottom: '20px', opacity: 0.9 }}>
        Selecciona tu modo de acceso
      </h2>

      <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', width: '100%', maxWidth: '300px' }}>
        {[
          { rol: 'operario', icon: '🏭', label: 'Operario', desc: 'Registro de producción en planta', color: '#F8A12F', colorDk: '#e6901e' },
          { rol: 'pcp', icon: '📋', label: 'PCP / Supervisor', desc: 'Validación y Kardex de MP', color: '#37BEEC', colorDk: '#0288d1' },
          { rol: 'gerencia', icon: '📊', label: 'Gerencia', desc: 'Panel de producción en tiempo real', color: '#7b1fa2', colorDk: '#4a148c' },
        ].map(({ rol, icon, label, desc, color, colorDk }) => (
          <button key={rol} onClick={() => seleccionarRol(rol)} style={{
            background: `linear-gradient(135deg, ${color}, ${colorDk})`,
            color: 'white', border: 'none',
            borderRadius: '14px', padding: '16px 20px',
            display: 'flex', alignItems: 'center', gap: '14px',
            textAlign: 'left', minHeight: '64px',
            boxShadow: `0 4px 16px ${color}40`,
          }}>
            <span style={{ fontSize: '24px', flexShrink: 0 }}>{icon}</span>
            <div>
              <div style={{ fontSize: '16px', fontWeight: 700 }}>{label}</div>
              <div style={{ fontSize: '11px', opacity: 0.8, marginTop: '2px' }}>{desc}</div>
            </div>
          </button>
        ))}
      </div>

      <div style={{ marginTop: '28px', display: 'flex', gap: '12px' }}>
        <button onClick={logout} style={{
          background: 'transparent', color: 'rgba(255,255,255,0.5)',
          border: '1px solid rgba(255,255,255,0.2)', borderRadius: '8px',
          padding: '8px 16px', fontSize: '12px',
        }}>Cerrar sesión</button>
      </div>

      <p style={{ color: 'rgba(255,255,255,0.25)', fontSize: '10px', textAlign: 'center', maxWidth: '260px', marginTop: '16px' }}>
        El acceso se configura automáticamente por rol en Maestro_Operarios
      </p>
    </div>
  )
}

// Router PCP
const GestionTablets = lazy(() => import('./pages/bi/GestionTablets'))

const RouterPCP = () => {
  const [pcpPantalla, setPcpPantalla] = useState('validacion')
  const { seleccionarRol, rolCuenta, rol } = useApp()
  const { logout } = useMsal()
  const esBI   = rolCuenta === 'bi'
  const esJefe = rol === 'jefeoperaciones'
  return (
    <Suspense fallback={<LoadingSpinner mensaje="Cargando..." />}>
      {pcpPantalla === 'validacion'
        ? <ValidacionPCP
            onIrKardex={() => setPcpPantalla('kardex')}
            onIrTablets={(esBI || esJefe) ? () => setPcpPantalla('tablets') : null}
            onLogout={logout}
            onCambiarRol={esBI ? () => seleccionarRol(null) : null} />
        : pcpPantalla === 'kardex'
          ? <KardexMP onVolver={() => setPcpPantalla('validacion')} onLogout={logout} />
          : <GestionTablets onVolver={() => setPcpPantalla('validacion')} onLogout={logout} />
      }
    </Suspense>
  )
}

// Roles que van a pantallas PCP/supervisión
const ROLES_PCP = ['pcp', 'bi', 'jefeoperaciones']

// Router Gerencia
const RouterGerencia = () => (
  <Suspense fallback={<LoadingSpinner mensaje="Cargando..." />}>
    <DashboardGerencia />
  </Suspense>
)

const Router = () => {
  const { pantalla, cargandoInicial, rol, seleccionarRol, setRolCuenta } = useApp()
  const { msalInstance, getToken, usuario } = useMsal()
  const [detectando, setDetectando] = useState(false)

  useSyncQueue({ msalInstance })

  // Detectar rol automáticamente desde Maestro_Operarios al iniciar sesión.
  // Solo corre si rol === null (evita sobreescribir una selección manual)
  useEffect(() => {
    if (rol !== null || !usuario?.email) return
    let cancelled = false
    const detectar = async () => {
      setDetectando(true)
      try {
        const token = await getToken()
        if (!token || cancelled) return
        const todos = await getListItems(token, 'Maestro_Operarios')
        if (cancelled) return
        const match = todos.find(o =>
          (o.Email || '').toLowerCase() === usuario.email.toLowerCase()
        )
        if (match?.Rol) {
          const rolNorm = match.Rol.toLowerCase().replace(/\s/g, '')
          setRolCuenta(rolNorm)
          seleccionarRol(rolNorm)
        }
        // Si no hay match → muestra SelectorRol (fallback manual)
      } catch {
        // Sin conexión o lista vacía → SelectorRol como fallback
      } finally {
        if (!cancelled) setDetectando(false)
      }
    }
    detectar()
    return () => { cancelled = true }
  }, [usuario?.email])

  if (cargandoInicial || detectando) return <LoadingSpinner mensaje="Verificando acceso..." />
  if (!rol) return <SelectorRol />
  if (rol === 'gerencia') return <RouterGerencia />
  if (ROLES_PCP.includes(rol)) return <RouterPCP />  // PCP/Jefe/BI → nunca necesitan seleccionar máquina

  // Solo operarios necesitan tablet configurada (permanente o temporal de sesión)
  if (!tabletConfigGuardada()) {
    return (
      <ConfigurarTablet
        onConfigurado={() => window.location.reload()}
        onTemporal={true}
      />
    )
  }

  const Page = {
    'bienvenida':      BienvenidaOperario,
    'inicio-turno':    InicioTurno,
    'turno-activo':    TurnoActivo,
    'parada':          ParadaMaquina,
    'cierre-turno':    CierreTurno,
    'opciones-turno':  OpcionesTurno,
    'cambio-producto': CambioProducto,
  }[pantalla] ?? BienvenidaOperario

  return (
    <Suspense fallback={<LoadingSpinner mensaje="Cargando pantalla..." />}>
      <Page />
    </Suspense>
  )
}

function App() {
  return (
    <>
      <UnauthenticatedTemplate>
        <Login />
      </UnauthenticatedTemplate>
      <AuthenticatedTemplate>
        <AppProvider>
          <Router />
        </AppProvider>
      </AuthenticatedTemplate>
    </>
  )
}

export default App
