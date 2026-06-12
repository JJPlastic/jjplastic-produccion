import { useState, useEffect } from 'react'
import { format, parseISO } from 'date-fns'
import { es } from 'date-fns/locale'
import { useMsal } from '../hooks/useMsal'
import { useApp } from '../context/AppContext'
import { Header } from '../components/Header'
import { LoadingSpinner } from '../components/LoadingSpinner'
import { getListItems, getOFsActivas, updateListItem } from '../services/sharepoint'
import { useMaestros } from '../hooks/useMaestros'
import { tabletConfig } from '../config/tabletConfig'

// eslint-disable-next-line no-unused-vars
const TURNO_LABEL = { M: 'Mañana', T: 'Tarde', N: 'Noche' }

export default function BienvenidaOperario() {
  const { getToken, usuario, logout } = useMsal()
  const { turnoActivo, setTurnoActivo, setPantalla, pendingCount, seleccionarRol, setModoRelevo, setOfPreseleccionada, setProductoPreseleccionado } = useApp()
  const { productos: catalogoProductos, materiasPrimas: catalogoMP } = useMaestros(getToken)

  const resolverNombre = (codigo) => {
    const catalogo = [...catalogoProductos, ...catalogoMP]
    const p = catalogo.find(x => (x.Codigo || '') === codigo || (x.Nombre || x.Title || '') === codigo)
    return p ? (p.Nombre || p.Title || codigo) : codigo
  }

  const [cargando, setCargando]               = useState(true)
  const [registrosActivos, setRegistrosActivos] = useState([])
  const [registrosCerrados, setRegistrosCerrados] = useState([])
  const [ofsKardex, setOfsKardex]             = useState([])
  const [cargandoReg, setCargandoReg]         = useState(null)
  const [pendingProductos, setPendingProductos] = useState([])
  const [kardexUltimaOF, setKardexUltimaOF]   = useState([]) // kardex items del último cierre

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
          const activosSP = deMiMaquina.filter(r => r.Estado === 'abierto' || r.Estado === 'transferido')

          // Auto-restaurar si SP tiene turno abierto pero IndexedDB lo perdió (cierre de sesión, etc.)
          if (activosSP.length > 0 && !turnoActivo) {
            const reg = activosSP[0]
            await setTurnoActivo({ ...reg, Paradas: reg.Paradas || '[]', spId: reg.ID })
            return // setTurnoActivo navega a turno-activo automáticamente
          }

          setRegistrosActivos(activosSP)
          const hoy = new Date().toISOString().split('T')[0]
          const cerradosHoy = deMiMaquina.filter(r =>
            r.Estado === 'cerrado' &&
            (r.HoraFin || r.Fecha || '').startsWith(hoy)
          ).sort((a, b) => new Date(b.HoraFin || 0) - new Date(a.HoraFin || 0))

          if (cerradosHoy.length > 0) {
            const ofMasReciente = cerradosHoy[0].Numero_OF
            const deEsaOF = ofMasReciente
              ? cerradosHoy.filter(r => r.Numero_OF === ofMasReciente)
              : [cerradosHoy[0]]
            setRegistrosCerrados(deEsaOF)

            // Cargar productos pendientes de esa OF
            if (ofMasReciente) {
              try {
                const kardexOf = await getListItems(token, 'Kardex_MP', {
                  filter: `Numero_OF eq '${ofMasReciente}'`, top: 100,
                })
                setKardexUltimaOF(kardexOf)
                const prodsOf = [...new Set(kardexOf.filter(k => k.Producto).map(k => k.Producto))]
                const completadosOF = new Set(deEsaOF.map(r => r.Producto))
                const activosOF = new Set(
                  deMiMaquina
                    .filter(r => r.Numero_OF === ofMasReciente && (r.Estado === 'abierto' || r.Estado === 'transferido'))
                    .map(r => r.Producto)
                )
                setPendingProductos(prodsOf.filter(p => !completadosOF.has(p) && !activosOF.has(p)))
              } catch { /* no crítico */ }
            }
          }
        }

        if (ofs.status === 'fulfilled') {
          const sinRegistro = ofs.value.filter(o =>
            o.tieneKardex && !o.tieneRegistro &&
            (!tabletConfig.codigoMaquina || o.maquina === tabletConfig.codigoMaquina)
          )
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
    await setTurnoActivo({
      ...registro,
      Paradas: registro.Paradas || '[]',
      spId: registro.ID,
      Estado: 'cerrado', // relevo: crea nuevo registro en cambio-producto
    })
    setModoRelevo(true)
    setPantalla('cambio-producto')
    setCargandoReg(null)
  }

  const cambiarProducto = async (registro) => {
    setCargandoReg(registro.ID)
    await setTurnoActivo({
      ...registro,
      Maquina: registro.Title || tabletConfig.codigoMaquina,
      Paradas: registro.Paradas || '[]',
      spId: registro.ID,
      Estado: 'cerrado',
    })
    setPantalla('cambio-producto')
    setCargandoReg(null)
  }

  if (cargando) return <LoadingSpinner mensaje="Cargando estado de la máquina..." />

  const fechaHoy = format(new Date(), "EEEE d 'de' MMMM", { locale: es })
  const ultimoReg = registrosCerrados[0] || null
  const primerNombre = usuario?.nombre?.split(' ')[0] || ''

  // Hay saldo si algún item del Kardex de la última OF todavía tiene kg disponibles
  const haySaldoMP = kardexUltimaOF.length > 0 && kardexUltimaOF.some(k => {
    const kg = (k.KgEntregados > 0) ? k.KgEntregados : (k.KgDeclaradoOperario || 0)
    const saldo = kg - (k.KgUsado || 0) - (k.KgMermaRec || 0) - (k.KgMermaNoRec || 0) - (k.KgDevueltos || 0)
    return saldo > 0.01
  })
  // Si el último turno fue cerrado con "se transfiere", la OF pasa a otra máquina
  const fueTransferida = !!(ultimoReg?.Transferencia_Pendiente)
  const puedeContinuar = haySaldoMP && !fueTransferida

  return (
    <div style={{
      background: 'linear-gradient(160deg, #003570 0%, #004895 55%, #005db0 100%)',
      minHeight: '100vh',
      display: 'flex',
      flexDirection: 'column',
    }}>
      <Header
        titulo="Bienvenido"
        subtitulo={fechaHoy}
        pendingCount={pendingCount}
        onLogout={logout}
        onCambiarRol={() => seleccionarRol(null)}
        color="#002d5a"
      />

      <div style={{
        padding: '16px 16px 32px',
        maxWidth: '480px',
        margin: '0 auto',
        width: '100%',
        flex: 1,
        display: 'flex',
        flexDirection: 'column',
        gap: '12px',
        justifyContent: 'center',
        boxSizing: 'border-box',
      }}>

        {/* Saludo */}
        {primerNombre && (
          <div style={{ textAlign: 'center', marginBottom: '4px' }}>
            <p style={{ fontSize: '24px', fontWeight: 800, color: 'white', margin: 0, letterSpacing: '-0.5px' }}>
              Hola, {primerNombre}
            </p>
            <p style={{ fontSize: '13px', color: 'rgba(255,255,255,0.55)', marginTop: '4px' }}>
              ¿Qué producción iniciamos hoy?
            </p>
          </div>
        )}

        {/* ── Producciones activas ── */}
        {gruposPorOf.map(grupo => grupo.items.map((reg) => (
          <div key={reg.ID} style={{
            backgroundColor: 'white',
            borderRadius: '18px',
            overflow: 'hidden',
            boxShadow: '0 6px 24px rgba(0,0,0,0.2)',
          }}>
            <div style={{
              background: 'linear-gradient(135deg, #37BEEC, #0288d1)',
              padding: '10px 16px',
              display: 'flex',
              alignItems: 'center',
              gap: '10px',
            }}>
              <span style={{
                width: '8px', height: '8px', borderRadius: '50%',
                backgroundColor: 'white', flexShrink: 0,
                boxShadow: '0 0 0 3px rgba(255,255,255,0.35)',
              }} />
              <span style={{ fontSize: '11px', color: 'white', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.08em' }}>
                {reg.Estado === 'transferido' ? '↔ Turno transferido' : '▶ En producción'}
              </span>
            </div>
            <div style={{ padding: '14px 16px 8px' }}>
              <p style={{ fontWeight: 800, fontSize: '16px', color: '#1a1a1a', margin: '0 0 8px', lineHeight: 1.3 }}>
                {resolverNombre(reg.Producto)}
              </p>
              <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
                {reg.Color && (
                  <span style={{ fontSize: '12px', color: '#555', background: '#f0f4ff', borderRadius: '6px', padding: '3px 10px' }}>
                    🎨 {reg.Color}
                  </span>
                )}
                <span style={{ fontSize: '12px', color: '#555', background: '#f0f4ff', borderRadius: '6px', padding: '3px 10px' }}>
                  👤 {reg.Operario}
                </span>
              </div>
            </div>
            <div style={{ padding: '8px 12px 14px' }}>
              <button
                onClick={async () => {
                  setCargandoReg(reg.ID)
                  await setTurnoActivo({ ...reg, Paradas: reg.Paradas || '[]', spId: reg.ID })
                  setCargandoReg(null)
                }}
                disabled={cargandoReg === reg.ID}
                style={{
                  width: '100%',
                  background: 'linear-gradient(135deg, #37BEEC, #0288d1)',
                  color: 'white', border: 'none', borderRadius: '10px',
                  padding: '14px', fontSize: '15px', fontWeight: 800,
                  cursor: 'pointer', minHeight: '50px',
                  boxShadow: '0 3px 10px rgba(55,190,236,0.4)',
                }}>
                {cargandoReg === reg.ID ? '⏳' : '▶ Volver al turno activo'}
              </button>
            </div>
          </div>
        )))}

        {/* ── Última OF ── */}
        {ultimoReg ? (
          <div style={{
            backgroundColor: 'white',
            borderRadius: '18px',
            overflow: 'hidden',
            boxShadow: '0 6px 24px rgba(0,0,0,0.2)',
          }}>
            <div style={{
              background: 'linear-gradient(135deg, #263238, #37474f)',
              padding: '10px 16px',
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
            }}>
              <span style={{ fontSize: '11px', color: 'rgba(255,255,255,0.65)', textTransform: 'uppercase', letterSpacing: '0.08em' }}>
                Última OF
              </span>
              <span style={{ fontFamily: 'monospace', fontSize: '10px', color: 'rgba(255,255,255,0.35)' }}>
                {ultimoReg.Numero_OF}
              </span>
            </div>
            <div style={{ padding: '14px 16px 10px' }}>
              <div style={{ display: 'flex', alignItems: 'flex-start', gap: '10px', marginBottom: '10px' }}>
                <span style={{ fontSize: '18px', flexShrink: 0, marginTop: '1px' }}>✅</span>
                <p style={{ fontWeight: 800, fontSize: '15px', color: '#1a1a1a', margin: 0, lineHeight: 1.3 }}>
                  {resolverNombre(ultimoReg.Producto)}
                </p>
              </div>
              <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
                {ultimoReg.Color && (
                  <span style={{ fontSize: '12px', color: '#555', background: '#f5f5f5', borderRadius: '6px', padding: '3px 9px' }}>
                    🎨 {ultimoReg.Color}
                  </span>
                )}
                <span style={{ fontSize: '12px', color: '#555', background: '#f5f5f5', borderRadius: '6px', padding: '3px 9px' }}>
                  👤 {ultimoReg.Operario}
                </span>
                {ultimoReg.HoraFin && (
                  <span style={{ fontSize: '12px', color: '#aaa', background: '#f5f5f5', borderRadius: '6px', padding: '3px 9px' }}>
                    🕐 {format(parseISO(ultimoReg.HoraFin), 'HH:mm')}
                  </span>
                )}
              </div>
            </div>
            <div style={{ padding: '4px 12px 14px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
              {/* Continuar mismo producto — solo si hay saldo de MP/MC en Kardex */}
              <button
                onClick={async () => {
                  setCargandoReg('relevo')
                  await setTurnoActivo({ ...ultimoReg, Paradas: ultimoReg.Paradas || '[]', spId: ultimoReg.ID, Estado: 'cerrado' })
                  setModoRelevo(true)
                  setPantalla('cambio-producto')
                  setCargandoReg(null)
                }}
                disabled={!!cargandoReg || !puedeContinuar}
                title={fueTransferida ? 'La OF fue transferida a otra máquina' : !haySaldoMP ? 'Sin saldo de MP disponible en Kardex para esta OF' : ''}
                style={{
                  width: '100%',
                  background: puedeContinuar
                    ? 'linear-gradient(135deg, #37BEEC, #0288d1)'
                    : '#bdbdbd',
                  color: 'white', border: 'none', borderRadius: '10px',
                  padding: '14px', fontSize: '14px', fontWeight: 700,
                  minHeight: '50px', cursor: puedeContinuar ? 'pointer' : 'not-allowed',
                  opacity: puedeContinuar ? 1 : 0.65,
                  display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px',
                }}>
                {cargandoReg === 'relevo' ? '⏳' : '🔄'} Continuar mismo producto
              </button>
              {!puedeContinuar && (
                <p style={{ fontSize: '11px', color: 'rgba(255,255,255,0.55)', margin: '0', textAlign: 'center' }}>
                  {fueTransferida ? 'OF transferida a otra máquina' : 'Sin saldo de MP disponible en Kardex'}
                </p>
              )}

              {/* Productos pendientes de la OF */}
              {pendingProductos.length > 0 ? (
                <>
                  <p style={{ fontSize: '11px', color: '#888', margin: '4px 0 0', textTransform: 'uppercase', letterSpacing: '0.05em', paddingLeft: '2px' }}>
                    Productos pendientes de la OF
                  </p>
                  {pendingProductos.map(cod => (
                    <button
                      key={cod}
                      onClick={async () => {
                        if (fueTransferida) return
                        setCargandoReg('prod-' + cod)
                        await setTurnoActivo({ ...ultimoReg, Paradas: ultimoReg.Paradas || '[]', spId: ultimoReg.ID, Estado: 'cerrado' })
                        setModoRelevo(false)
                        setProductoPreseleccionado(cod)
                        setPantalla('cambio-producto')
                        setCargandoReg(null)
                      }}
                      disabled={!!cargandoReg || fueTransferida}
                      title={fueTransferida ? 'OF transferida a otra máquina' : ''}
                      style={{
                        width: '100%',
                        background: fueTransferida ? '#bdbdbd' : 'linear-gradient(135deg, #004895, #1565c0)',
                        color: 'white', border: 'none', borderRadius: '10px',
                        padding: '14px', fontSize: '14px', fontWeight: 700,
                        minHeight: '50px', cursor: fueTransferida ? 'not-allowed' : 'pointer',
                        opacity: fueTransferida ? 0.65 : 1,
                        boxShadow: fueTransferida ? 'none' : '0 3px 10px rgba(0,72,149,0.3)',
                        display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px',
                      }}>
                      {cargandoReg === 'prod-' + cod ? '⏳' : '▶'} {resolverNombre(cod)}
                    </button>
                  ))}
                  {fueTransferida && (
                    <p style={{ fontSize: '11px', color: 'rgba(255,255,255,0.55)', margin: '0', textAlign: 'center' }}>
                      OF transferida a otra máquina
                    </p>
                  )}
                </>
              ) : !ultimoReg?.Numero_OF ? (
                /* "Otro producto" — solo si el turno anterior fue sin OF de PCP */
                <button
                  onClick={async () => {
                    setCargandoReg('cambio')
                    await setTurnoActivo({ ...ultimoReg, Paradas: ultimoReg.Paradas || '[]', spId: ultimoReg.ID, Estado: 'cerrado' })
                    setModoRelevo(false)
                    setPantalla('cambio-producto')
                    setCargandoReg(null)
                  }}
                  disabled={!!cargandoReg}
                  style={{
                    width: '100%',
                    background: 'linear-gradient(135deg, #004895, #1565c0)',
                    color: 'white', border: 'none', borderRadius: '10px',
                    padding: '14px', fontSize: '14px', fontWeight: 700,
                    minHeight: '50px', cursor: 'pointer',
                    boxShadow: '0 3px 10px rgba(0,72,149,0.3)',
                    display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px',
                  }}>
                  {cargandoReg === 'cambio' ? '⏳' : '📦'} Otro producto
                </button>
              ) : null}
            </div>
          </div>
        ) : null}

        {/* ── OF preparada por PCP ── */}
        {ofsKardex.length > 0 ? ofsKardex.map(of => (
          <div key={of.of} style={{
            backgroundColor: 'white',
            borderRadius: '18px',
            overflow: 'hidden',
            boxShadow: '0 6px 24px rgba(0,0,0,0.2)',
          }}>
            <div style={{
              background: 'linear-gradient(135deg, #2e7d32, #43a047)',
              padding: '10px 16px',
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
            }}>
              <span style={{ fontSize: '11px', color: 'white', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.08em' }}>
                📦 MP preparada por PCP
              </span>
              <span style={{ fontSize: '10px', color: 'rgba(255,255,255,0.55)', fontFamily: 'monospace' }}>{of.of}</span>
            </div>
            <div style={{ padding: '10px 16px 0' }}>
              {of.mpPorProducto && Object.entries(of.mpPorProducto).map(([prod, mps], pIdx) => (
                <div key={prod || pIdx} style={{ marginBottom: '8px' }}>
                  {/* Sub-cabecera de producto */}
                  <div style={{
                    backgroundColor: '#f0fff4',
                    borderLeft: '3px solid #2e7d32',
                    padding: '5px 10px',
                    marginBottom: '4px',
                    borderRadius: '0 6px 6px 0',
                  }}>
                    <p style={{ fontSize: '13px', fontWeight: 800, color: '#2e7d32', margin: 0 }}>
                      {prod ? resolverNombre(prod) : '—'}
                    </p>
                  </div>
                  {/* MPs de este producto */}
                  {mps.map((m, i) => (
                    <div key={i} style={{
                      display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                      padding: '5px 4px',
                      borderBottom: i < mps.length - 1 ? '1px solid #f0f0f0' : 'none',
                    }}>
                      <span style={{ fontSize: '13px', color: '#555' }}>• {resolverNombre(m.insumo)}</span>
                      <span style={{ fontSize: '13px', color: '#1a1a1a', fontWeight: 700 }}>
                        {(m.kg - m.kgDev).toFixed(2)} kg
                      </span>
                    </div>
                  ))}
                </div>
              ))}
            </div>
            <div style={{ padding: '12px' }}>
              <button
                onClick={() => { setOfPreseleccionada(of); setPantalla('inicio-turno') }}
                style={{
                  width: '100%', padding: '14px',
                  background: 'linear-gradient(135deg, #2e7d32, #43a047)',
                  color: 'white', border: 'none', borderRadius: '10px',
                  fontSize: '15px', fontWeight: 800, cursor: 'pointer',
                  boxShadow: '0 3px 10px rgba(46,125,50,0.35)',
                }}>
                ▶ Iniciar turno con OF
              </button>
            </div>
          </div>
        )) : (
          <div style={{
            backgroundColor: 'rgba(255,255,255,0.1)',
            borderRadius: '16px',
            padding: '16px',
            border: '1.5px solid rgba(255,255,255,0.18)',
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '14px' }}>
              <span style={{ fontSize: '28px', flexShrink: 0 }}>📦</span>
              <div>
                <p style={{ fontSize: '11px', color: 'rgba(255,255,255,0.45)', textTransform: 'uppercase', letterSpacing: '0.05em', margin: '0 0 3px' }}>
                  MP preparada por PCP
                </p>
                <p style={{ fontSize: '13px', color: 'rgba(255,255,255,0.8)', margin: 0 }}>
                  PCP aún no registró MP para esta máquina
                </p>
              </div>
            </div>
            <button
              onClick={() => setPantalla('inicio-turno')}
              style={{
                width: '100%', padding: '13px',
                backgroundColor: 'white', color: '#004895',
                border: 'none', borderRadius: '10px',
                fontSize: '14px', fontWeight: 800, cursor: 'pointer',
              }}>
              ▶ Iniciar turno sin OF de PCP
            </button>
          </div>
        )}


      </div>
    </div>
  )
}
