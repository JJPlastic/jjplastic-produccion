import { useState, useEffect, useCallback, useRef } from 'react'
import { useMsal } from '../../hooks/useMsal'
import { useApp } from '../../context/AppContext'
import { getListItems } from '../../services/sharepoint'
import { MAQUINAS_DISPONIBLES } from '../../config/tabletConfig'
import { format, formatDistanceToNow } from 'date-fns'
import { es } from 'date-fns/locale'

const INTERVALO_SEG = 60

const TURNO_LABEL = { M: 'Mañana', T: 'Tarde', N: 'Noche' }
const COLOR_ESTADO = { Verde: '#4caf50', Amarillo: '#ff9800', Rojo: '#f44336' }

// ── Tarjeta de máquina ────────────────────────────────────────────────────────
const TarjetaMaquina = ({ maq, resolverProd }) => {
  const { codigoMaquina, nombreMaquina, registros, ultimoParcial } = maq

  // Registro abierto (puede haber 0 o 1)
  const abierto   = registros.find(r => r.Estado === 'abierto') || null
  const activa    = !!abierto
  const enParada  = activa && !!(abierto?.En_Parada)

  // Unidades del turno activo (del parcial) o del registro al cierre
  const confTurno = ultimoParcial?.Acum_Conf_Turno ?? (abierto?.UnidadesConformes ?? 0)
  const defTurno  = ultimoParcial?.Acum_Def_Turno  ?? (abierto?.UnidadesDefectuosas ?? 0)

  // Acumulado del lote (varios turnos del mismo producto)
  const confLote  = ultimoParcial?.Acum_Conf_Lote ?? confTurno

  // Tiempo sin reporte de parcial
  const msDesdeUltimoParcial = ultimoParcial
    ? Date.now() - new Date(ultimoParcial.Timestamp).getTime()
    : null
  const minSinReporte = msDesdeUltimoParcial != null
    ? Math.floor(msDesdeUltimoParcial / 60000)
    : null
  const sinReporte = activa && !enParada && minSinReporte !== null && minSinReporte > 120

  // Registros cerrados hoy (para historial rápido)
  const cerradosHoy = registros.filter(r => r.Estado === 'cerrado')

  // Colores de borde según estado
  const borderColor = !activa ? '#e0e0e0'
    : enParada ? '#c62828'
    : sinReporte ? '#ff9800'
    : abierto?.Estado_Validacion ? (COLOR_ESTADO[abierto.Estado_Validacion] || '#4caf50')
    : '#4caf50'

  const bgHead = !activa ? '#f5f5f5'
    : enParada ? '#ffebee'
    : sinReporte ? '#fff3e0'
    : '#f0fff4'

  const badgeColor = !activa ? '#9e9e9e'
    : enParada ? '#c62828'
    : sinReporte ? '#ff9800'
    : '#4caf50'

  const badgeLabel = !activa ? 'Sin turno'
    : enParada ? 'En parada'
    : sinReporte ? 'Sin reporte'
    : 'Activa'

  return (
    <div style={{
      backgroundColor: 'white',
      borderRadius: '14px',
      border: `2px solid ${borderColor}`,
      overflow: 'hidden',
      opacity: activa ? 1 : 0.6,
      boxShadow: activa ? '0 2px 10px rgba(0,0,0,0.07)' : 'none',
      transition: 'border-color 0.3s',
    }}>
      {/* Cabecera */}
      <div style={{
        backgroundColor: bgHead,
        padding: '10px 14px',
        borderBottom: `1px solid ${borderColor}30`,
        display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start',
      }}>
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <strong style={{ fontSize: '17px', color: '#1a1a1a' }}>{codigoMaquina}</strong>
            <span style={{
              backgroundColor: badgeColor, color: 'white',
              borderRadius: '20px', padding: '2px 9px',
              fontSize: '10px', fontWeight: 700, textTransform: 'uppercase',
            }}>{badgeLabel}</span>
            {abierto?.Estado_Validacion && (
              <span style={{
                width: '10px', height: '10px', borderRadius: '50%',
                backgroundColor: COLOR_ESTADO[abierto.Estado_Validacion] || '#ccc',
                display: 'inline-block', flexShrink: 0,
              }} title={`Semáforo: ${abierto.Estado_Validacion}`} />
            )}
          </div>
          <p style={{ fontSize: '11px', color: '#888', marginTop: '2px' }}>{nombreMaquina}</p>
        </div>
        {activa && abierto?.Turno && (
          <span style={{
            fontSize: '11px', color: '#555',
            backgroundColor: 'rgba(0,0,0,0.07)', borderRadius: '6px', padding: '3px 8px',
            whiteSpace: 'nowrap',
          }}>
            {TURNO_LABEL[abierto.Turno] || abierto.Turno}
          </span>
        )}
      </div>

      {activa ? (
        <div style={{ padding: '12px 14px', display: 'flex', flexDirection: 'column', gap: '10px' }}>

          {/* Producto + color + OF */}
          <div>
            <p style={{ fontSize: '15px', fontWeight: 800, color: '#1a1a1a', lineHeight: 1.3, margin: 0 }}>
              {resolverProd(abierto.Producto)}
            </p>
            {abierto.Color && (
              <p style={{ fontSize: '12px', color: '#555', margin: '2px 0 0' }}>🎨 {abierto.Color}</p>
            )}
            {abierto.Numero_OF && (
              <p style={{ fontSize: '10px', color: '#aaa', fontFamily: 'monospace', margin: '2px 0 0' }}>
                {abierto.Numero_OF}
              </p>
            )}
          </div>

          {/* Métricas */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
            <div style={{
              backgroundColor: '#f0fff4', borderRadius: '10px',
              padding: '10px', textAlign: 'center',
            }}>
              <div style={{ fontSize: '26px', fontWeight: 900, color: '#2e7d32', lineHeight: 1 }}>
                {confTurno.toLocaleString()}
              </div>
              <div style={{ fontSize: '10px', color: '#555', textTransform: 'uppercase', marginTop: '3px' }}>
                Conformes turno
              </div>
            </div>
            <div style={{
              backgroundColor: defTurno > 0 ? '#fff3e0' : '#fafafa', borderRadius: '10px',
              padding: '10px', textAlign: 'center',
            }}>
              <div style={{ fontSize: '26px', fontWeight: 900, color: defTurno > 0 ? '#e65100' : '#ccc', lineHeight: 1 }}>
                {defTurno.toLocaleString()}
              </div>
              <div style={{ fontSize: '10px', color: '#555', textTransform: 'uppercase', marginTop: '3px' }}>
                Defectuosas
              </div>
            </div>
          </div>

          {/* Acumulado lote (si difiere del turno) */}
          {confLote > confTurno && (
            <div style={{
              backgroundColor: '#e8f0fb', borderRadius: '8px',
              padding: '6px 10px',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
            }}>
              <span style={{ fontSize: '11px', color: '#004895', fontWeight: 600 }}>Acum. lote</span>
              <span style={{ fontSize: '15px', fontWeight: 800, color: '#004895' }}>
                {confLote.toLocaleString()} und.
              </span>
            </div>
          )}

          {/* Operario + inicio */}
          <div style={{
            display: 'flex', justifyContent: 'space-between', alignItems: 'center',
            borderTop: '1px solid #f0f0f0', paddingTop: '8px',
          }}>
            <span style={{ fontSize: '12px', color: '#444' }}>
              👤 {abierto.Operario || '—'}
            </span>
            <span style={{ fontSize: '11px', color: '#aaa' }}>
              Inicio {abierto.HoraInicio ? format(new Date(abierto.HoraInicio), 'HH:mm') : '—'}
            </span>
          </div>

          {/* Banner parada activa */}
          {enParada && (
            <div style={{
              backgroundColor: '#c62828', color: 'white',
              borderRadius: '8px', padding: '8px 12px',
              display: 'flex', alignItems: 'center', gap: '8px',
            }}>
              <span style={{ fontSize: '18px' }}>⏸</span>
              <div>
                <div style={{ fontSize: '13px', fontWeight: 800 }}>Máquina en parada</div>
                {(() => {
                  const paradas = (() => { try { return JSON.parse(abierto.Paradas || '[]') } catch { return [] } })()
                  const minAcum = paradas.reduce((s, p) => s + (p.duracion_minutos || 0), 0)
                  return (
                    <div style={{ fontSize: '11px', opacity: 0.85 }}>
                      {paradas.length} parada{paradas.length !== 1 ? 's' : ''} registrada{paradas.length !== 1 ? 's' : ''}
                      {minAcum > 0 && ` · ${Math.round(minAcum)} min acum.`}
                    </div>
                  )
                })()}
              </div>
            </div>
          )}

          {/* Último parcial */}
          {!enParada && ultimoParcial ? (
            <div style={{
              fontSize: '11px',
              color: sinReporte ? '#e65100' : '#bbb',
              display: 'flex', justifyContent: 'space-between',
            }}>
              <span>{sinReporte ? '⚠ Sin reporte:' : 'Último reporte:'}</span>
              <span>
                {formatDistanceToNow(new Date(ultimoParcial.Timestamp), { locale: es, addSuffix: false })} atrás
              </span>
            </div>
          ) : !enParada && (
            <p style={{ fontSize: '11px', color: '#ccc', textAlign: 'center', margin: 0 }}>
              Sin parciales registrados aún
            </p>
          )}
        </div>
      ) : (
        /* Sin turno activo — mostrar resumen si hubo turnos hoy */
        <div style={{ padding: cerradosHoy.length > 0 ? '12px 14px' : '28px 14px', textAlign: 'center' }}>
          {cerradosHoy.length > 0 ? (
            <div>
              <p style={{ fontSize: '11px', color: '#888', marginBottom: '8px', fontWeight: 600 }}>
                Turnos cerrados hoy
              </p>
              {cerradosHoy.map(r => (
                <div key={r.ID} style={{
                  padding: '6px 0', borderBottom: '1px solid #f5f5f5',
                }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '12px' }}>
                    <span style={{ color: '#333', fontWeight: 600 }}>{resolverProd(r.Producto)}</span>
                    <span style={{ color: '#2e7d32', fontWeight: 700 }}>
                      {(r.UnidadesConformes || 0).toLocaleString()} ✓
                    </span>
                  </div>
                  {r.Color && (
                    <span style={{ fontSize: '11px', color: '#888' }}>🎨 {r.Color}</span>
                  )}
                </div>
              ))}
            </div>
          ) : (
            <div>
              <p style={{ fontSize: '28px' }}>💤</p>
              <p style={{ fontSize: '12px', color: '#bbb' }}>Sin actividad hoy</p>
            </div>
          )}
        </div>
      )}
    </div>
  )
}

// ── Dashboard principal ───────────────────────────────────────────────────────
export default function DashboardGerencia() {
  const { getToken, logout } = useMsal()
  const { seleccionarRol }  = useApp()

  const [datos, setDatos]                     = useState([])
  const [cargando, setCargando]               = useState(true)
  const [ultimaAct, setUltimaAct]             = useState(null)
  const [countdown, setCountdown]             = useState(INTERVALO_SEG)
  const [nombresProductos, setNombresProductos] = useState({})

  // Cargar nombres de productos una vez
  useEffect(() => {
    let cancelled = false
    const cargar = async () => {
      try {
        const token = await getToken()
        if (!token) return
        const todos = await getListItems(token, 'Maestro_Productos', { top: 500 })
        if (cancelled) return
        const map = {}
        todos.forEach(p => {
          const codigo = p.Codigo || p.Nombre || p.Title || ''
          const nombre = p.Nombre || p.Title || codigo
          if (codigo) map[codigo] = nombre
        })
        setNombresProductos(map)
      } catch {}
    }
    cargar()
    return () => { cancelled = true }
  }, [])

  const resolverProd = useCallback((codigo) =>
    nombresProductos[codigo] || codigo || '—'
  , [nombresProductos])

  const cargarDatos = useCallback(async () => {
    try {
      const token = await getToken()
      if (!token) return

      const hoy = new Date()
      hoy.setHours(0, 0, 0, 0)
      const hoyISO = hoy.toISOString().split('T')[0]

      // Todos los registros de producción de hoy (abiertos y cerrados)
      const registros = await getListItems(token, 'Registro_Produccion', {
        filter: `Fecha ge '${hoyISO}T00:00:00Z'`,
        top: 100,
        orderby: 'HoraInicio desc',
      })

      // Parciales de las últimas 24 h
      const desde24h = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString()
      let parciales = []
      try {
        parciales = await getListItems(token, 'Registro_Produccion_Parcial', {
          filter: `Timestamp ge '${desde24h}'`,
          orderby: 'Timestamp desc',
          top: 500,
        })
      } catch { /* lista puede estar vacía */ }

      // Último parcial por Registro_ID (ya vienen ordenados desc → el primero es el último)
      const ultimoPorReg = {}
      parciales.forEach(p => {
        if (!ultimoPorReg[p.Registro_ID]) ultimoPorReg[p.Registro_ID] = p
      })

      // Agrupar registros por máquina
      const porMaquina = {}
      registros.forEach(r => {
        const maq = r.Title
        if (!porMaquina[maq]) porMaquina[maq] = []
        porMaquina[maq].push(r)
      })

      // Construir array de máquinas con su registro abierto y último parcial
      const resultado = MAQUINAS_DISPONIBLES.map(m => {
        const regsDeEsta = porMaquina[m.codigoMaquina] || []
        const abierto = regsDeEsta.find(r => r.Estado === 'abierto') || null
        const ultimoParcial = abierto ? (ultimoPorReg[abierto.ID] || null) : null
        return {
          ...m,
          registros: regsDeEsta,
          ultimoParcial,
        }
      })

      setDatos(resultado)
      setUltimaAct(new Date())
    } catch (err) {
      console.error('Error dashboard gerencia:', err)
    } finally {
      setCargando(false)
    }
  }, [getToken])

  // Carga inicial + polling cada INTERVALO_SEG
  useEffect(() => {
    cargarDatos()
    const pollTimer = setInterval(() => {
      cargarDatos()
      setCountdown(INTERVALO_SEG)
    }, INTERVALO_SEG * 1000)
    const tickTimer = setInterval(() => setCountdown(p => Math.max(0, p - 1)), 1000)
    return () => { clearInterval(pollTimer); clearInterval(tickTimer) }
  }, [cargarDatos])

  const handleRefresh = () => { cargarDatos(); setCountdown(INTERVALO_SEG) }

  // KPIs globales
  const enParadaCount = datos.filter(d => {
    const ab = d.registros.find(r => r.Estado === 'abierto')
    return !!ab?.En_Parada
  }).length
  const activas  = datos.filter(d => {
    const ab = d.registros.find(r => r.Estado === 'abierto')
    return ab && !ab.En_Parada
  }).length
  const sinTurno = MAQUINAS_DISPONIBLES.length - activas - enParadaCount
  const totalConf   = datos.reduce((s, d) => {
    const up = d.ultimoParcial
    if (up) return s + (up.Acum_Conf_Turno || 0)
    const ab = d.registros.find(r => r.Estado === 'abierto')
    return s + (ab?.UnidadesConformes || 0)
  }, 0)
  const totalCerradosHoy = datos.reduce((s, d) =>
    s + d.registros.filter(r => r.Estado === 'cerrado').reduce((ss, r) => ss + (r.UnidadesConformes || 0), 0)
  , 0)

  return (
    <div style={{ backgroundColor: '#f0f2f5', minHeight: '100vh' }}>

      {/* Header */}
      <header style={{
        background: 'linear-gradient(135deg, #37006e 0%, #6a1b9a 100%)',
        color: 'white', padding: '10px 16px',
        position: 'sticky', top: 0, zIndex: 100,
        boxShadow: '0 2px 8px rgba(0,0,0,0.2)',
      }}>
        <div style={{ maxWidth: '1200px', margin: '0 auto', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '12px' }}>
          <div>
            <p style={{ fontSize: '10px', opacity: 0.65, textTransform: 'uppercase', letterSpacing: '0.1em', margin: 0 }}>
              Panel Gerencia
            </p>
            <h1 style={{ fontSize: '17px', fontWeight: 800, margin: 0 }}>Producción en tiempo real</h1>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <div style={{ textAlign: 'right', fontSize: '11px', opacity: 0.7, lineHeight: 1.4 }}>
              {ultimaAct && <p style={{ margin: 0 }}>Actualizado {format(ultimaAct, 'HH:mm:ss')}</p>}
              <p style={{ margin: 0 }}>↺ en {countdown}s</p>
            </div>
            <button
              onClick={handleRefresh}
              style={{
                backgroundColor: 'rgba(255,255,255,0.2)', color: 'white',
                border: '1.5px solid rgba(255,255,255,0.45)', borderRadius: '8px',
                padding: '6px 14px', fontSize: '13px', fontWeight: 700, cursor: 'pointer',
              }}>
              Actualizar
            </button>
            <button
              onClick={() => seleccionarRol(null)}
              style={{
                backgroundColor: 'transparent', color: 'rgba(255,255,255,0.65)',
                border: '1px solid rgba(255,255,255,0.3)', borderRadius: '8px',
                padding: '6px 10px', fontSize: '11px', cursor: 'pointer',
              }}>
              Salir
            </button>
          </div>
        </div>
      </header>

      <div style={{ padding: '14px 12px', maxWidth: '1200px', margin: '0 auto' }}>

        {/* KPIs resumen */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '10px', marginBottom: '16px' }}>
          {[
            { valor: activas,                          label: 'Produciendo',              color: '#4caf50' },
            { valor: enParadaCount,                    label: 'En parada',                color: '#c62828' },
            { valor: sinTurno,                         label: 'Sin turno',                color: '#9e9e9e' },
            { valor: totalConf.toLocaleString(),       label: 'Und. conformes activas',   color: '#004895' },
          ].map(({ valor, label, color }) => (
            <div key={label} style={{
              backgroundColor: 'white', borderRadius: '12px',
              padding: '14px 12px', textAlign: 'center',
              boxShadow: '0 1px 4px rgba(0,0,0,0.06)',
              borderTop: `3px solid ${color}`,
            }}>
              <div style={{ fontSize: '26px', fontWeight: 900, color, lineHeight: 1 }}>{valor}</div>
              <div style={{ fontSize: '11px', color: '#666', marginTop: '4px' }}>{label}</div>
            </div>
          ))}
        </div>

        {/* Leyenda */}
        <div style={{ display: 'flex', gap: '16px', marginBottom: '12px', flexWrap: 'wrap' }}>
          {[
            { color: '#4caf50', label: 'Produciendo' },
            { color: '#c62828', label: 'En parada' },
            { color: '#ff9800', label: 'Sin reporte +2h' },
            { color: '#e0e0e0', label: 'Sin turno' },
          ].map(({ color, label }) => (
            <div key={label} style={{ display: 'flex', alignItems: 'center', gap: '5px', fontSize: '11px', color: '#666' }}>
              <span style={{ width: '10px', height: '10px', borderRadius: '50%', backgroundColor: color, display: 'inline-block' }} />
              {label}
            </div>
          ))}
          <div style={{ marginLeft: 'auto', fontSize: '11px', color: '#aaa' }}>
            Refresca automáticamente cada {INTERVALO_SEG}s
          </div>
        </div>

        {/* Grid de máquinas */}
        {cargando ? (
          <div style={{ textAlign: 'center', padding: '60px', color: '#888', fontSize: '14px' }}>
            Cargando datos de planta...
          </div>
        ) : (
          <div style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))',
            gap: '12px',
          }}>
            {datos.map(maq => (
              <TarjetaMaquina
                key={maq.codigoMaquina}
                maq={maq}
                resolverProd={resolverProd}
              />
            ))}
          </div>
        )}
      </div>
    </div>
  )
}
