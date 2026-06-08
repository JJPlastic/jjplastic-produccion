import { useState, useEffect } from 'react'
import { format, parseISO } from 'date-fns'
import { es } from 'date-fns/locale'
import { useMsal } from '../hooks/useMsal'
import { useApp } from '../context/AppContext'
import { Header } from '../components/Header'
import { LoadingSpinner } from '../components/LoadingSpinner'
import { getListItems, getOFsActivas } from '../services/sharepoint'
import { useMaestros } from '../hooks/useMaestros'
import { tabletConfig } from '../config/tabletConfig'

// eslint-disable-next-line no-unused-vars
const TURNO_LABEL = { M: 'Mañana', T: 'Tarde', N: 'Noche' }

export default function BienvenidaOperario() {
  const { getToken, usuario, logout } = useMsal()
  const { setTurnoActivo, setPantalla, pendingCount, seleccionarRol, setModoRelevo, setOfPreseleccionada } = useApp()
  const { productos: catalogoProductos } = useMaestros(getToken)

  const resolverNombre = (codigo) => {
    const p = catalogoProductos.find(x => (x.Codigo || '') === codigo || (x.Nombre || x.Title || '') === codigo)
    return p ? (p.Nombre || p.Title || codigo) : codigo
  }

  const [cargando, setCargando]               = useState(true)
  const [registrosActivos, setRegistrosActivos] = useState([]) // abiertos/transferidos
  const [registrosCerrados, setRegistrosCerrados] = useState([]) // cerrados hoy
  const [ofsKardex, setOfsKardex]             = useState([])  // MP entregada sin registro
  const [cargandoReg, setCargandoReg]         = useState(null)

  useEffect(() => {
    let cancelled = false
    const cargar = async () => {
      try {
        const token = await getToken()
        if (!token) return

        const [regs, ofs] = await Promise.allSettled([
          getListItems(token, 'Registro_Produccion', { top: 100 }),
          getOFsActivas(token, tabletConfig.codigoMaquina),
        ])

        if (cancelled) return

        if (regs.status === 'fulfilled') {
          const deMiMaquina = regs.value.filter(r => (r.Title || '') === tabletConfig.codigoMaquina)
          setRegistrosActivos(deMiMaquina.filter(r =>
            r.Estado === 'abierto' || r.Estado === 'transferido'
          ))
          // Registros cerrados HOY — solo la OF más reciente (agrupada)
          const hoy = new Date().toISOString().split('T')[0]
          const cerradosHoy = deMiMaquina.filter(r =>
            r.Estado === 'cerrado' &&
            (r.HoraFin || r.Fecha || '').startsWith(hoy)
          ).sort((a, b) => new Date(b.HoraFin || 0) - new Date(a.HoraFin || 0))

          if (cerradosHoy.length > 0) {
            // Tomar la OF del registro más reciente
            const ofMasReciente = cerradosHoy[0].Numero_OF
            // Todos los registros de ESA OF
            const deEsaOF = ofMasReciente
              ? cerradosHoy.filter(r => r.Numero_OF === ofMasReciente)
              : [cerradosHoy[0]]
            setRegistrosCerrados(deEsaOF)
          }
        }

        if (ofs.status === 'fulfilled') {
          // OFs con Kardex pero SIN registro activo (PCP preparó, operario no inició)
          const sinRegistro = ofs.value.filter(o => o.tieneKardex && !o.tieneRegistro)
          setOfsKardex(sinRegistro)
        }
      } catch (err) {
        console.error('Error cargando bienvenida:', err)
      } finally {
        if (!cancelled) setCargando(false)
      }
    }
    if (navigator.onLine) cargar()
    else setCargando(false)
    return () => { cancelled = true }
  }, [])

  // Agrupar registros activos por OF
  const gruposPorOf = (() => {
    const map = {}
    registrosActivos.forEach(r => {
      const of = r.Numero_OF || 'sin-of'
      if (!map[of]) map[of] = { numeroOF: r.Numero_OF, turno: r.Turno, items: [] }
      map[of].items.push(r)
    })
    return Object.values(map)
  })()

  const continuarProducto = async (registro) => {
    setCargandoReg(registro.ID)
    // Cargar el registro como turnoActivo
    await setTurnoActivo({
      Maquina: registro.Title || tabletConfig.codigoMaquina,
      Turno: registro.Turno,
      Operario: registro.Operario,
      Operario_Apoyo: registro.Operario_Apoyo,
      Producto: registro.Producto,
      ProductoNombre: registro.Producto, // Se mostrará el código hasta que se cargue el nombre
      Color: registro.Color,
      HoraInicio: registro.HoraInicio,
      Fecha: registro.Fecha,
      Estado: registro.Estado,
      Tipo_Apertura: registro.Tipo_Apertura,
      Registro_Previo_ID: registro.Registro_Previo_ID,
      Paradas: registro.Paradas || '[]',
      Estado_Validacion: registro.Estado_Validacion,
      Numero_OF: registro.Numero_OF,
      Codigo_Lote: registro.Codigo_Lote,
      KgPorUnidadProducto: 0,
      spId: registro.ID,
      localId: registro.ID?.toString(),
    })
    setCargandoReg(null)
  }

  const cambiarProducto = async (registro) => {
    setCargandoReg(registro.ID)
    await setTurnoActivo({
      ...registro,
      Maquina: registro.Title || tabletConfig.codigoMaquina,
      Paradas: registro.Paradas || '[]',
      spId: registro.ID,
      Estado: 'cerrado', // para que CambioProducto sepa que viene de aquí
    })
    setPantalla('cambio-producto')
    setCargandoReg(null)
  }

  if (cargando) return <LoadingSpinner mensaje="Cargando estado de la máquina..." />

  const fechaHoy = format(new Date(), "EEEE d 'de' MMMM", { locale: es })

  const hayActividad = gruposPorOf.length > 0 || ofsKardex.length > 0 || registrosCerrados.length > 0

  const ultimoReg = registrosCerrados[0] || null

  return (
    <div style={{ backgroundColor: '#f0f2f5', minHeight: '100vh', display: 'flex', flexDirection: 'column' }}>
      <Header titulo="Bienvenido" subtitulo={fechaHoy} pendingCount={pendingCount}
        onLogout={logout} onCambiarRol={() => seleccionarRol(null)} color="#004895" />

      <div style={{ padding: '12px 14px', maxWidth: '480px', margin: '0 auto', width: '100%', flex: 1, display: 'flex', flexDirection: 'column', gap: '10px', justifyContent: 'center' }}>

        {/* ── Producciones activas ── */}
        {gruposPorOf.map(grupo => grupo.items.map((reg, idx) => (
          <div key={reg.ID} style={{ backgroundColor: '#004895', borderRadius: '14px', padding: '14px 16px', color: 'white', boxShadow: '0 3px 12px rgba(0,72,149,0.2)' }}>
            <p style={{ fontSize: '10px', opacity: 0.7, textTransform: 'uppercase', letterSpacing: '0.05em', margin: '0 0 4px' }}>▶ En producción</p>
            <p style={{ fontWeight: 800, fontSize: '16px', margin: '0 0 3px' }}>{resolverNombre(reg.Producto)}</p>
            <p style={{ fontSize: '12px', opacity: 0.8, margin: '0 0 10px' }}>
              {reg.Color && <span>🎨 {reg.Color} · </span>}
              <span>👤 {reg.Operario}</span>
              {reg.Estado === 'transferido' && <span style={{ color: '#ffcc80', marginLeft: '6px' }}>↔ Transferido</span>}
            </p>
            <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: '8px' }}>
              <button onClick={() => continuarProducto(reg)} disabled={cargandoReg === reg.ID}
                style={{ backgroundColor: 'white', color: '#004895', border: 'none', borderRadius: '8px', padding: '12px', fontSize: '14px', fontWeight: 800, cursor: 'pointer', minHeight: '48px' }}>
                {cargandoReg === reg.ID ? '⏳' : '▶ Continuar'}
              </button>
              <button onClick={() => cambiarProducto(reg)} disabled={cargandoReg === reg.ID}
                style={{ backgroundColor: 'rgba(255,255,255,0.12)', color: 'white', border: '1.5px solid rgba(255,255,255,0.35)', borderRadius: '8px', padding: '12px', fontSize: '13px', fontWeight: 600, cursor: 'pointer', minHeight: '48px' }}>
                Cambiar
              </button>
            </div>
          </div>
        )))}

        {/* ── OF preparada por PCP ── */}
        {ofsKardex.length > 0 ? ofsKardex.map(of => (
          <div key={of.of} style={{ background: 'linear-gradient(135deg,#1b5e20,#2e7d32)', borderRadius: '14px', padding: '16px', boxShadow: '0 4px 16px rgba(27,94,32,0.2)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '8px' }}>
              <p style={{ fontSize: '10px', color: 'rgba(255,255,255,0.7)', textTransform: 'uppercase', letterSpacing: '0.05em', margin: 0 }}>📦 MP preparada por PCP</p>
              <span style={{ fontSize: '10px', color: 'rgba(255,255,255,0.45)', fontFamily: 'monospace' }}>{of.of}</span>
            </div>
            {of.resumenMP && of.resumenMP.split(' · ').map((item, i) => {
              const partes = item.split(': ')
              const nombre = partes[0] || item
              const kg = partes[1] || ''
              return (
                <div key={i} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2px' }}>
                  <span style={{ fontSize: '12px', color: 'rgba(255,255,255,0.9)' }}>• {nombre}</span>
                  <span style={{ fontSize: '12px', color: 'white', fontWeight: 700, flexShrink: 0, marginLeft: '8px' }}>{kg}</span>
                </div>
              )
            })}
            <button onClick={() => { setOfPreseleccionada(of); setPantalla('inicio-turno') }} style={{
              marginTop: '12px', width: '100%', padding: '14px',
              backgroundColor: 'white', color: '#1b5e20', border: 'none', borderRadius: '10px', fontSize: '15px', fontWeight: 800, cursor: 'pointer',
            }}>▶ Iniciar turno con OF</button>
          </div>
        )) : (
          <div style={{ borderRadius: '14px', border: '1.5px dashed #c8e6c9', backgroundColor: 'white', overflow: 'hidden' }}>
            <div style={{ padding: '14px 16px' }}>
              <p style={{ fontSize: '10px', color: '#aaa', textTransform: 'uppercase', letterSpacing: '0.05em', margin: '0 0 4px' }}>📦 MP preparada por PCP</p>
              <p style={{ fontSize: '13px', color: '#888', margin: 0 }}>PCP aún no ha registrado MP para esta máquina.</p>
            </div>
            <div style={{ borderTop: '1px solid #f0f0f0', padding: '12px 16px', backgroundColor: '#fafafa' }}>
              <p style={{ fontSize: '11px', color: '#aaa', margin: '0 0 8px' }}>Puedes declarar la MP recibida manualmente:</p>
              <button onClick={() => setPantalla('inicio-turno')} style={{
                width: '100%', padding: '12px', backgroundColor: '#004895', color: 'white',
                border: 'none', borderRadius: '8px', fontSize: '14px', fontWeight: 700, cursor: 'pointer',
              }}>▶ Iniciar turno sin OF de PCP</button>
            </div>
          </div>
        )}

        {/* ── Última OF ── */}
        {ultimoReg ? (
          <div style={{ backgroundColor: 'white', borderRadius: '14px', overflow: 'hidden', border: '1.5px solid #e0e0e0', boxShadow: '0 2px 8px rgba(0,0,0,0.07)' }}>
            {/* Header */}
            <div style={{ background: 'linear-gradient(135deg,#263238,#37474f)', color: 'white', padding: '10px 14px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span style={{ fontSize: '10px', opacity: 0.7, textTransform: 'uppercase', letterSpacing: '0.08em' }}>Última OF</span>
              <span style={{ fontFamily: 'monospace', fontSize: '10px', opacity: 0.45 }}>{ultimoReg.Numero_OF}</span>
            </div>
            {/* Info producto */}
            <div style={{ padding: '12px 14px 10px' }}>
              <p style={{ fontWeight: 800, fontSize: '15px', color: '#1a1a1a', margin: '0 0 4px' }}>
                <span style={{ color: '#2e7d32', marginRight: '5px' }}>✓</span>
                {resolverNombre(ultimoReg.Producto)}
              </p>
              <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
                {ultimoReg.Color && <span style={{ fontSize: '12px', color: '#555', background: '#f5f5f5', borderRadius: '6px', padding: '2px 8px' }}>🎨 {ultimoReg.Color}</span>}
                <span style={{ fontSize: '12px', color: '#555' }}>👤 {ultimoReg.Operario}</span>
                {ultimoReg.HoraFin && <span style={{ fontSize: '11px', color: '#aaa', marginLeft: 'auto' }}>{format(parseISO(ultimoReg.HoraFin), 'HH:mm')}</span>}
              </div>
            </div>
            {/* Botones acción */}
            <div style={{ padding: '0 12px 12px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
              <button
                onClick={async () => { setCargandoReg('relevo'); await setTurnoActivo({ ...ultimoReg, Paradas: ultimoReg.Paradas || '[]', spId: ultimoReg.ID, Estado: 'cerrado' }); setModoRelevo(true); setPantalla('cambio-producto'); setCargandoReg(null) }}
                disabled={!!cargandoReg}
                style={{ width: '100%', background: 'linear-gradient(135deg,#0288d1,#37BEEC)', color: 'white', border: 'none', borderRadius: '10px', padding: '14px', fontSize: '14px', fontWeight: 700, minHeight: '50px', cursor: 'pointer', boxShadow: '0 2px 8px rgba(55,190,236,0.35)', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }}>
                {cargandoReg === 'relevo' ? '⏳' : '🔄'} Continuar mismo producto
              </button>
              <button
                onClick={async () => { setCargandoReg('cambio'); await setTurnoActivo({ ...ultimoReg, Paradas: ultimoReg.Paradas || '[]', spId: ultimoReg.ID, Estado: 'cerrado' }); setModoRelevo(false); setPantalla('cambio-producto'); setCargandoReg(null) }}
                disabled={!!cargandoReg}
                style={{ width: '100%', background: 'linear-gradient(135deg,#004895,#1565c0)', color: 'white', border: 'none', borderRadius: '10px', padding: '14px', fontSize: '14px', fontWeight: 700, minHeight: '50px', cursor: 'pointer', boxShadow: '0 2px 8px rgba(0,72,149,0.3)', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }}>
                {cargandoReg === 'cambio' ? '⏳' : '📦'} Otro producto
              </button>
            </div>
          </div>
        ) : (
          <div style={{ backgroundColor: 'white', borderRadius: '14px', border: '1.5px dashed #e0e0e0', overflow: 'hidden' }}>
            <div style={{ backgroundColor: '#f5f5f5', padding: '8px 14px' }}>
              <span style={{ fontSize: '11px', color: '#bbb', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Última OF</span>
            </div>
            <div style={{ padding: '14px 16px', textAlign: 'center' }}>
              <p style={{ fontSize: '13px', color: '#ccc', margin: '0 0 10px' }}>Aún no se produjo nada en esta máquina hoy.</p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                <button disabled style={{ padding: '13px', borderRadius: '10px', backgroundColor: '#f5f5f5', color: '#ccc', border: 'none', fontSize: '13px', fontWeight: 700, cursor: 'not-allowed' }}>🔄 Continuar mismo producto</button>
                <button disabled style={{ padding: '13px', borderRadius: '10px', backgroundColor: '#f5f5f5', color: '#ccc', border: 'none', fontSize: '13px', fontWeight: 700, cursor: 'not-allowed' }}>📦 Otro producto</button>
              </div>
            </div>
          </div>
        )}

      </div>
    </div>
  )
}
