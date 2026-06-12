import { useState, useEffect, useRef } from 'react'
import { format, differenceInSeconds, parseISO } from 'date-fns'
import { es } from 'date-fns/locale'
import { useMsal } from '../hooks/useMsal'
import { useApp } from '../context/AppContext'
import { useWakeLock } from '../hooks/useWakeLock'
import { Header } from '../components/Header'
import { uploadAttachment, createListItem, updateListItem, getListItem, getListItems } from '../services/sharepoint'

import { encolarOperacion } from '../services/indexedDB'
import { useMaestros } from '../hooks/useMaestros'
import { SearchSelect } from '../components/SearchSelect'
import { useProduccionParcial } from '../hooks/useProduccionParcial'
import { Toast, mensajeRed } from '../components/Toast'

// Formatea segundos → HH:MM:SS
const formatCronometro = (secs) => {
  const h = Math.floor(secs / 3600)
  const m = Math.floor((secs % 3600) / 60)
  const s = secs % 60
  return [h, m, s].map((v) => String(v).padStart(2, '0')).join(':')
}

const KpiCard = ({ label, valor, color = '#004895', unidad = '' }) => (
  <div style={{
    backgroundColor: 'white',
    borderRadius: '14px',
    padding: '16px',
    boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: '4px',
  }}>
    <span style={{ fontSize: '11px', color: '#888', textTransform: 'uppercase', letterSpacing: '0.05em', textAlign: 'center' }}>{label}</span>
    <span style={{ fontSize: '28px', fontWeight: 800, color, lineHeight: 1 }}>{valor}</span>
    {unidad && <span style={{ fontSize: '11px', color: '#888' }}>{unidad}</span>}
  </div>
)

const BigButton = ({ onClick, color, children, disabled }) => (
  <button
    onClick={onClick}
    disabled={disabled}
    style={{
      backgroundColor: disabled ? '#ccc' : color,
      color: 'white',
      border: 'none',
      borderRadius: '14px',
      padding: '18px 12px',
      fontSize: '16px',
      fontWeight: 700,
      minHeight: '64px',
      width: '100%',
      cursor: disabled ? 'not-allowed' : 'pointer',
      boxShadow: '0 3px 10px rgba(0,0,0,0.15)',
      lineHeight: 1.3,
    }}
  >
    {children}
  </button>
)

export default function TurnoActivo() {
  const { getToken, logout } = useMsal()
  const { turnoActivo, actualizarTurnoLocal, pendingCount, setPantalla, limpiarTurno, seleccionarRol } = useApp()
  const [segundos, setSegundos]         = useState(0)
  const [subiendo, setSubiendo]           = useState(false)
  const [modalParcial, setModalParcial]   = useState(false)
  const [modalMP, setModalMP]             = useState(false)
  const [mpFilas, setMpFilas]             = useState([{ id: crypto.randomUUID(), insumo: '', kg: '' }])
  const [guardandoMP, setGuardandoMP]     = useState(false)
  const [toast, setToast]                 = useState(null)

  const { materiasPrimas, productos: catalogoProductos } = useMaestros(getToken)
  const [parcConf, setParcConf]         = useState('')
  const [parcDef, setParcDef]           = useState('')
  const [parcObs, setParcObs]           = useState('')
  const [guardandoParcial, setGuardandoParcial] = useState(false)
  const fotoInputRef = useRef(null)

  const [minutosSinReporte, setMinutosSinReporte] = useState(0)
  const [acumOF, setAcumOF] = useState(0) // total conformes cerrados en esta OF

  // Cargar total de unidades conformes de registros CERRADOS del mismo lote
  useEffect(() => {
    if (!turnoActivo?.Codigo_Lote || !navigator.onLine) return
    let cancelled = false
    const cargar = async () => {
      try {
        const token = await getToken()
        if (!token) return
        const regs = await getListItems(token, 'Registro_Produccion', { top: 200 })
        if (cancelled) return
        const total = regs
          .filter(r => r.Codigo_Lote === turnoActivo.Codigo_Lote &&
                       r.Estado === 'cerrado' &&
                       r.ID !== turnoActivo.spId)
          .reduce((s, r) => s + (r.UnidadesConformes || 0), 0)
        setAcumOF(total)
      } catch { /* offline */ }
    }
    cargar()
    return () => { cancelled = true }
  }, [turnoActivo?.Codigo_Lote])

  const { acumConfTurno, acumDefTurno, acumConfLote, ultimoTimestamp, agregarParcial } = useProduccionParcial({
    getToken,
    registroId: turnoActivo?.spId,
    codigoLote: turnoActivo?.Codigo_Lote,
  })

  useWakeLock(true)

  // Recordatorio cada 2 horas sin reporte parcial
  useEffect(() => {
    const calcular = () => {
      const referencia = ultimoTimestamp || turnoActivo?.HoraInicio
      if (!referencia) return
      const mins = Math.floor((Date.now() - new Date(referencia).getTime()) / 60000)
      setMinutosSinReporte(mins)
    }
    calcular()
    const t = setInterval(calcular, 60000)
    return () => clearInterval(t)
  }, [ultimoTimestamp, turnoActivo?.HoraInicio])

  // Verificar que el registro aún existe en SP (por si fue eliminado manualmente)
  useEffect(() => {
    if (!turnoActivo?.spId || !navigator.onLine) return
    const verificar = async () => {
      try {
        const token = await getToken()
        if (!token) return
        const item = await getListItem(token, 'Registro_Produccion', turnoActivo.spId)
        const estado = item?.Estado || ''
        if (!['abierto', 'transferido'].includes(estado)) {
          await limpiarTurno()
        }
      } catch {
        // Si 404 u otro error, el registro ya no existe
        await limpiarTurno()
      }
    }
    verificar()
  }, [turnoActivo?.spId])

  // Cronómetro desde HoraInicio
  useEffect(() => {
    const calcular = () => {
      const inicio = parseISO(turnoActivo.HoraInicio)
      setSegundos(differenceInSeconds(new Date(), inicio))
    }
    calcular()
    const timer = setInterval(calcular, 1000)
    return () => clearInterval(timer)
  }, [turnoActivo?.HoraInicio])

  // Extraer KPIs del turno activo (actualizados en cierre)
  const paradas = (() => {
    try { return JSON.parse(turnoActivo?.Paradas || '[]') }
    catch { return [] }
  })()

  const minsParadas = paradas.reduce((acc, p) => acc + (p.duracion_minutos || 0), 0)

  const handleFoto = async (e) => {
    const file = e.target.files?.[0]
    if (!file) return
    setSubiendo(true)
    try {
      if (turnoActivo.spId) {
        const token = await getToken()
        if (token) {
          const nombre = `foto_turno_${turnoActivo.spId}_${Date.now()}.jpg`
          await uploadAttachment(token, 'Registro_Produccion', turnoActivo.spId, nombre, file)
          await updateListItem(token, 'Registro_Produccion', turnoActivo.spId, { TieneFoto: true })
        }
      }
      setToast({ mensaje: '✓ Foto adjuntada correctamente.', tipo: 'ok' })
    } catch (err) {
      setToast({ mensaje: mensajeRed(err), tipo: 'error' })
    } finally {
      setSubiendo(false)
      e.target.value = ''
    }
  }

  // Resuelve código o nombre → objeto MP del maestro
  const findMP = (val) => materiasPrimas.find(p =>
    (p.Codigo || '') === val ||
    (p.Nombre || p.Title || '').toLowerCase() === (val || '').toLowerCase()
  )

  const handleGuardarMP = async () => {
    const validas = mpFilas.filter(f => f.insumo && parseFloat(f.kg) > 0)
    if (!validas.length) return
    setGuardandoMP(true)

    // Helper: encolar cada fila como create para sync posterior
    const encolarFilas = (obs) => Promise.allSettled(validas.map(f => {
      const insumoCodigo = findMP(f.insumo)?.Codigo || f.insumo
      return encolarOperacion({
        tipo: 'create',
        listName: 'Kardex_MP',
        data: {
          Title: turnoActivo.Maquina || '',
          Fecha: new Date().toISOString(),
          Turno: turnoActivo.Turno,
          Insumo: insumoCodigo,
          KgEntregados: 0,
          KgDeclaradoOperario: parseFloat(f.kg),
          KgDevueltos: 0,
          Observacion: obs,
          Numero_OF: turnoActivo.Numero_OF || '',
          Producto: turnoActivo.Producto,
        },
      })
    }))

    try {
      const token = await getToken()
      if (token && navigator.onLine) {
        const kardexActual = await getListItems(token, 'Kardex_MP', {
          filter: `Numero_OF eq '${turnoActivo.Numero_OF || ''}'`, top: 100,
        })
        await Promise.allSettled(validas.map(f => {
          const kgNuevo = parseFloat(f.kg)
          const mpObj = findMP(f.insumo)
          const insumoCodigo = mpObj?.Codigo || f.insumo
          const existente = kardexActual.find(k => {
            const mpK = findMP(k.Insumo)
            if (mpObj && mpK) return (mpObj.Codigo || mpObj.ID) === (mpK.Codigo || mpK.ID)
            return (k.Insumo || '').trim().toLowerCase() === (f.insumo || '').trim().toLowerCase()
          })
          if (existente) {
            const kgActual = existente.KgDeclaradoOperario ?? existente.KgEntregados ?? 0
            return updateListItem(token, 'Kardex_MP', existente.ID, {
              Insumo: insumoCodigo,
              KgDeclaradoOperario: parseFloat((kgActual + kgNuevo).toFixed(3)),
              Observacion: 'Actualizado por operario — pendiente validación PCP',
            })
          } else {
            return createListItem(token, 'Kardex_MP', {
              Title: turnoActivo.Maquina || '',
              Fecha: new Date().toISOString(),
              Turno: turnoActivo.Turno,
              Insumo: insumoCodigo,
              KgEntregados: 0,
              KgDeclaradoOperario: kgNuevo,
              KgDevueltos: 0,
              Observacion: 'Declarado por operario — pendiente validación PCP',
              Numero_OF: turnoActivo.Numero_OF || '',
              Producto: turnoActivo.Producto,
            })
          }
        }))
        setToast({ mensaje: '✓ MP guardada en Kardex.', tipo: 'ok' })
      } else {
        await encolarFilas('Declarado por operario (sin red) — pendiente validación PCP')
        setToast({ mensaje: 'Sin red — MP guardada localmente, se enviará al reconectar.', tipo: 'warn' })
      }
      setModalMP(false)
      setMpFilas([{ id: crypto.randomUUID(), insumo: '', kg: '' }])
    } catch (err) {
      // Red disponible pero falló — encolar igual
      await encolarFilas('Declarado por operario (error red) — pendiente validación PCP')
      setModalMP(false)
      setMpFilas([{ id: crypto.randomUUID(), insumo: '', kg: '' }])
      setToast({ mensaje: 'Sin red — MP guardada localmente, se enviará al reconectar.', tipo: 'warn' })
    } finally {
      setGuardandoMP(false)
    }
  }

  const handleGuardarParcial = async () => {
    const conf = parseInt(parcConf) || 0
    const def  = parseInt(parcDef)  || 0
    if (conf === 0 && def === 0) return
    setGuardandoParcial(true)
    try {
      const resultado = await agregarParcial({
        undConformes:   conf,
        undDefectuosas: def,
        operario:       turnoActivo.Operario,
        obs:            parcObs,
      })
      setModalParcial(false)
      setParcConf(''); setParcDef(''); setParcObs('')
      if (resultado?._offline) {
        setToast({ mensaje: 'Sin red — parcial guardado localmente, se enviará al reconectar.', tipo: 'warn' })
      }
    } catch (err) {
      setToast({ mensaje: mensajeRed(err), tipo: 'error' })
    } finally {
      setGuardandoParcial(false)
    }
  }

  if (!turnoActivo) return null

  const turnoLabels = { M: 'Mañana', T: 'Tarde', N: 'Noche' }

  return (
    <>
    <div style={{ backgroundColor: '#f5f5f5', minHeight: '100vh' }}>
      <Header
        titulo="Turno activo"
        subtitulo={`${turnoLabels[turnoActivo.Turno] || turnoActivo.Turno} · ${format(parseISO(turnoActivo.HoraInicio), 'HH:mm', { locale: es })}`}
        pendingCount={pendingCount}
        onLogout={logout}
        onCambiarRol={() => seleccionarRol(null)}
        color="#2e7d32"
      />

      <div style={{ padding: '16px', maxWidth: '540px', margin: '0 auto', display: 'flex', flexDirection: 'column', gap: '14px' }}>

        {/* ── Tarjeta de producto ── */}
        {(() => {
          const prod = catalogoProductos.find(p =>
            (p.Codigo || '') === turnoActivo.Producto || (p.Nombre || p.Title || '') === turnoActivo.Producto
          )
          const nombre = prod ? (prod.Nombre || prod.Title || turnoActivo.Producto) : turnoActivo.Producto
          const pctDef = acumConfTurno + acumDefTurno > 0 ? (acumDefTurno / (acumConfTurno + acumDefTurno)) * 100 : 0
          return (
            <div style={{
              background: 'linear-gradient(135deg, #1b5e20 0%, #2e7d32 100%)',
              borderRadius: '16px', padding: '14px 16px', color: 'white',
              boxShadow: '0 4px 16px rgba(46,125,50,0.3)',
            }}>
              <p style={{ fontWeight: 800, fontSize: '18px', margin: 0, lineHeight: 1.15, letterSpacing: '-0.01em' }}>{nombre}</p>
              <div style={{ display: 'flex', alignItems: 'center', gap: '14px', marginTop: '8px' }}>
                {turnoActivo.Color && (
                  <span style={{ display: 'flex', alignItems: 'center', gap: '4px', fontSize: '13px', backgroundColor: 'rgba(255,255,255,0.15)', borderRadius: '20px', padding: '3px 10px' }}>
                    🎨 {turnoActivo.Color}
                  </span>
                )}
                <span style={{ display: 'flex', alignItems: 'center', gap: '4px', fontSize: '13px', backgroundColor: 'rgba(255,255,255,0.15)', borderRadius: '20px', padding: '3px 10px' }}>
                  👤 {turnoActivo.Operario}
                </span>
                {pctDef > 5 && (
                  <span style={{ marginLeft: 'auto', fontSize: '11px', backgroundColor: '#d32f2f', borderRadius: '12px', padding: '2px 8px', fontWeight: 700 }}>
                    ⚠ {pctDef.toFixed(0)}% def.
                  </span>
                )}
              </div>
            </div>
          )
        })()}

        {/* ── KPIs ── */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px' }}>

          {/* Cronómetro */}
          <div style={{ backgroundColor: 'white', borderRadius: '14px', padding: '16px', boxShadow: '0 2px 8px rgba(0,0,0,0.07)', textAlign: 'center', gridColumn: '1' }}>
            <p style={{ fontSize: '10px', color: '#aaa', textTransform: 'uppercase', letterSpacing: '0.07em', margin: 0 }}>Cronómetro</p>
            <p style={{ fontSize: '30px', fontWeight: 800, color: '#2e7d32', margin: '6px 0 0', fontVariantNumeric: 'tabular-nums', letterSpacing: '-0.02em' }}>{formatCronometro(segundos)}</p>
          </div>

          {/* Conformes */}
          <div style={{ backgroundColor: 'white', borderRadius: '14px', padding: '16px', boxShadow: '0 2px 8px rgba(0,0,0,0.07)', textAlign: 'center' }}>
            <p style={{ fontSize: '10px', color: '#aaa', textTransform: 'uppercase', letterSpacing: '0.07em', margin: 0 }}>Conformes</p>
            <p style={{ fontSize: '36px', fontWeight: 800, color: '#2e7d32', margin: '6px 0 0', lineHeight: 1 }}>
              {acumConfTurno > 0 ? acumConfTurno.toLocaleString() : '—'}
            </p>
            {acumConfLote > 0 && acumConfLote !== acumConfTurno && (
              <p style={{ fontSize: '10px', color: '#aaa', margin: '4px 0 0' }}>{acumConfLote.toLocaleString()} en lote</p>
            )}
          </div>

          {/* Paradas */}
          <div style={{ backgroundColor: 'white', borderRadius: '14px', padding: '14px 16px', boxShadow: '0 2px 8px rgba(0,0,0,0.07)', textAlign: 'center' }}>
            <p style={{ fontSize: '10px', color: '#aaa', textTransform: 'uppercase', letterSpacing: '0.07em', margin: 0 }}>Paradas</p>
            <p style={{ fontSize: '32px', fontWeight: 800, color: paradas.length > 0 ? '#c62828' : '#bbb', margin: '6px 0 0', lineHeight: 1 }}>{paradas.length}</p>
            {minsParadas > 0 && <p style={{ fontSize: '10px', color: '#aaa', margin: '4px 0 0' }}>{minsParadas} min</p>}
          </div>

          {/* Defectuosas */}
          <div style={{ backgroundColor: 'white', borderRadius: '14px', padding: '14px 16px', boxShadow: '0 2px 8px rgba(0,0,0,0.07)', textAlign: 'center' }}>
            <p style={{ fontSize: '10px', color: '#aaa', textTransform: 'uppercase', letterSpacing: '0.07em', margin: 0 }}>Defectuosas</p>
            <p style={{ fontSize: '32px', fontWeight: 800, color: acumDefTurno > 0 ? '#F8A12F' : '#bbb', margin: '6px 0 0', lineHeight: 1 }}>
              {acumDefTurno > 0 ? acumDefTurno : '—'}
            </p>
          </div>
        </div>

        {/* Acumulado del lote — suma registros cerrados del mismo lote + avances actuales */}
        {turnoActivo.Codigo_Lote && (
          <div style={{ backgroundColor: '#e8f5e9', borderRadius: '10px', padding: '8px 14px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: '11px', color: '#2e7d32', fontFamily: 'monospace', fontWeight: 600 }}>{turnoActivo.Numero_OF}</span>
            <div style={{ textAlign: 'right' }}>
              <strong style={{ fontSize: '13px', color: '#1b5e20' }}>
                {(acumOF + acumConfTurno).toLocaleString()} und. en lote
              </strong>
              {acumOF > 0 && acumConfTurno > 0 && (
                <p style={{ fontSize: '10px', color: '#888', margin: 0 }}>
                  {acumOF} prev. + {acumConfTurno} este turno
                </p>
              )}
            </div>
          </div>
        )}

        {/* ── Botones de acción ── */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginTop: '4px' }}>
          {/* Reporte */}
          <button onClick={() => setModalParcial(true)} style={{
            backgroundColor: '#004895', color: 'white', border: 'none', borderRadius: '14px',
            padding: '18px', fontSize: '16px', fontWeight: 700, minHeight: '60px', width: '100%', cursor: 'pointer',
          }}>📋 Reporte de avance</button>

          {/* Parada + Finalizar */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px' }}>
            <button onClick={async () => {
              setPantalla('parada')
              // Marcar en SP para que Gerencia vea la parada en tiempo real
              await actualizarTurnoLocal({ En_Parada: true })
              if (turnoActivo?.spId && navigator.onLine) {
                try {
                  const token = await getToken()
                  if (token) updateListItem(token, 'Registro_Produccion', turnoActivo.spId, { En_Parada: true })
                } catch {}
              }
            }} style={{
              backgroundColor: '#c62828', color: 'white', border: 'none', borderRadius: '14px',
              padding: '16px', fontSize: '15px', fontWeight: 700, minHeight: '56px', cursor: 'pointer',
            }}>⏸ Parada</button>
            <button onClick={() => setPantalla('cierre-turno')} style={{
              backgroundColor: '#004895', color: 'white', border: 'none', borderRadius: '14px',
              padding: '16px', fontSize: '15px', fontWeight: 700, minHeight: '56px', cursor: 'pointer',
            }}>🏁 Finalizar</button>
          </div>

          {/* Agregar MP + Foto — secundarios */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px' }}>
            <button onClick={() => setModalMP(true)} style={{
              backgroundColor: 'white', color: '#004895', border: '2px solid #004895', borderRadius: '12px',
              padding: '13px', fontSize: '13px', fontWeight: 600, minHeight: '48px', cursor: 'pointer',
            }}>📦 Agregar MP</button>
            <button onClick={() => fotoInputRef.current?.click()} disabled={subiendo} style={{
              backgroundColor: 'white', color: '#555', border: '2px solid #ddd', borderRadius: '12px',
              padding: '13px', fontSize: '13px', fontWeight: 600, minHeight: '48px', cursor: subiendo ? 'not-allowed' : 'pointer',
            }}>{subiendo ? '⏳...' : '📷 Foto'}</button>
          </div>
        </div>

        {/* Modal Agregar MP → Kardex */}
        {modalMP && (
          <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 300, padding: '20px' }}>
            <div style={{ backgroundColor: 'white', borderRadius: '16px', padding: '22px', width: '100%', maxWidth: '380px', display: 'flex', flexDirection: 'column', gap: '14px' }}>
              <h3 style={{ color: '#004895', fontSize: '16px', fontWeight: 700 }}>📦 Registrar MP recibida</h3>
              <p style={{ fontSize: '12px', color: '#666' }}>
                Se guardará en Kardex vinculada a la OF del turno.
              </p>

              {/* Filas de MP */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                {mpFilas.map((fila, idx) => (
                  <div key={fila.id} style={{ display: 'grid', gridTemplateColumns: '1fr 90px 32px', gap: '8px', alignItems: 'center' }}>
                    <SearchSelect
                      opciones={materiasPrimas.map(mp => ({
                        value: mp.Codigo || mp.Nombre || mp.Title || '',
                        label: mp.Nombre || mp.Title || '',
                      }))}
                      value={fila.insumo}
                      onChange={v => setMpFilas(p => p.map(f => f.id === fila.id ? { ...f, insumo: v } : f))}
                      placeholder="Insumo..."
                    />
                    <input type="number" step="0.01" min="0"
                      value={fila.kg || ''}
                      onChange={e => setMpFilas(p => p.map(f => f.id === fila.id ? { ...f, kg: e.target.value } : f))}
                      placeholder="kg"
                      style={{ padding: '10px', borderRadius: '8px', border: '2px solid #ddd', fontSize: '15px', fontWeight: 700, textAlign: 'right', color: '#1a1a1a', backgroundColor: 'white' }}
                    />
                    <button type="button" onClick={() => setMpFilas(p => p.length > 1 ? p.filter(f => f.id !== fila.id) : p)}
                      style={{ background: mpFilas.length === 1 ? '#f5f5f5' : '#ffebee', color: mpFilas.length === 1 ? '#ccc' : '#c62828', border: 'none', borderRadius: '6px', width: '32px', height: '32px', fontSize: '16px', cursor: 'pointer' }}>×</button>
                  </div>
                ))}
                <button type="button" onClick={() => setMpFilas(p => [...p, { id: crypto.randomUUID(), insumo: '', kg: '' }])}
                  style={{ background: 'transparent', color: '#004895', border: '1.5px solid #004895', borderRadius: '8px', padding: '7px', fontSize: '13px', fontWeight: 700, cursor: 'pointer' }}>
                  + Agregar insumo
                </button>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', gap: '10px', marginTop: '4px' }}>
                <button onClick={() => { setModalMP(false); setMpFilas([{ id: crypto.randomUUID(), insumo: '', kg: '' }]) }}
                  style={{ padding: '13px', borderRadius: '10px', border: '2px solid #ddd', background: 'white', color: '#555', fontSize: '14px', cursor: 'pointer' }}>
                  Cancelar
                </button>
                <button onClick={handleGuardarMP} disabled={guardandoMP}
                  style={{ padding: '13px', borderRadius: '10px', border: 'none', background: guardandoMP ? '#ccc' : '#004895', color: 'white', fontSize: '15px', fontWeight: 700, cursor: guardandoMP ? 'not-allowed' : 'pointer' }}>
                  {guardandoMP ? '⏳ Guardando...' : '✓ Guardar en Kardex'}
                </button>
              </div>
            </div>
          </div>
        )}


        {/* Modal registro parcial de producción */}
        {modalParcial && (
          <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 300, padding: '24px' }}>
            <div style={{ backgroundColor: 'white', borderRadius: '16px', padding: '24px', width: '100%', maxWidth: '360px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
              <h3 style={{ color: '#004895', fontSize: '16px', fontWeight: 700 }}>📊 Registrar producción</h3>

              {/* Acumulado actual */}
              <div style={{ backgroundColor: '#e8f5e9', borderRadius: '8px', padding: '10px 14px', display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px', fontSize: '13px' }}>
                <div><span style={{ color: '#555' }}>Acum. conformes: </span><strong style={{ color: '#2e7d32' }}>{acumConfTurno}</strong></div>
                <div><span style={{ color: '#555' }}>Acum. defectuosas: </span><strong style={{ color: '#c62828' }}>{acumDefTurno}</strong></div>
              </div>

              <div>
                <label style={{ fontSize: '13px', fontWeight: 600, color: '#333', display: 'block', marginBottom: '6px' }}>Unidades conformes (este conteo) *</label>
                <input type="number" inputMode="numeric" min="0" value={parcConf}
                  onChange={e => setParcConf(e.target.value)} autoFocus placeholder="0"
                  style={{ width: '100%', padding: '14px', borderRadius: '10px', border: '2px solid #004895', fontSize: '22px', fontWeight: 800, textAlign: 'right', color: '#1a1a1a', backgroundColor: 'white' }}
                />
                {parseInt(parcConf) > 0 && (
                  <p style={{ fontSize: '12px', color: '#2e7d32', marginTop: '4px', fontWeight: 600 }}>
                    Nuevo acumulado: {acumConfTurno + parseInt(parcConf)} und
                  </p>
                )}
              </div>

              <div>
                <label style={{ fontSize: '13px', fontWeight: 600, color: '#333', display: 'block', marginBottom: '6px' }}>Unidades defectuosas (este conteo)</label>
                <input type="number" inputMode="numeric" min="0" value={parcDef}
                  onChange={e => setParcDef(e.target.value)} placeholder="0"
                  style={{ width: '100%', padding: '12px', borderRadius: '10px', border: '2px solid #ddd', fontSize: '18px', fontWeight: 700, textAlign: 'right', color: '#1a1a1a', backgroundColor: 'white' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '13px', fontWeight: 600, color: '#333', display: 'block', marginBottom: '6px' }}>Observación (opcional)</label>
                <input value={parcObs} onChange={e => setParcObs(e.target.value)}
                  placeholder="Ej: Cambio de molde, reinicio de máquina..."
                  style={{ width: '100%', padding: '10px', borderRadius: '8px', border: '2px solid #ddd', fontSize: '14px', color: '#1a1a1a', backgroundColor: 'white' }}
                />
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px' }}>
                <button onClick={() => { setModalParcial(false); setParcConf(''); setParcDef(''); setParcObs('') }} style={{
                  padding: '14px', borderRadius: '10px', border: '2px solid #ccc',
                  backgroundColor: 'white', color: '#555', fontSize: '15px', cursor: 'pointer',
                }}>Cancelar</button>
                <button onClick={handleGuardarParcial} disabled={guardandoParcial || (!parseInt(parcConf) && !parseInt(parcDef))} style={{
                  padding: '14px', borderRadius: '10px', border: 'none',
                  backgroundColor: guardandoParcial || (!parseInt(parcConf) && !parseInt(parcDef)) ? '#ccc' : '#004895',
                  color: 'white', fontSize: '15px', fontWeight: 700,
                  cursor: guardandoParcial ? 'not-allowed' : 'pointer',
                }}>{guardandoParcial ? '⏳...' : '✓ Guardar'}</button>
              </div>
            </div>
          </div>
        )}

        {/* Input oculto para cámara */}
        <input
          ref={fotoInputRef}
          type="file"
          accept="image/*"
          capture="environment"
          style={{ display: 'none' }}
          onChange={handleFoto}
        />

        {/* Estado offline */}
        {!turnoActivo.spId && (
          <div style={{
            backgroundColor: '#fff3e0',
            border: '1px solid #F8A12F',
            borderRadius: '10px',
            padding: '10px 14px',
            fontSize: '13px',
            color: '#e65100',
          }}>
            ⚠ Registro guardado localmente — se sincronizará cuando haya conexión.
          </div>
        )}
      </div>
    </div>

    <Toast toast={toast} onClose={() => setToast(null)} />

    {/* Overlay alerta 2h sin reporte */}
    {minutosSinReporte >= 120 && (
      <div style={{
        position: 'fixed', inset: 0, zIndex: 9000,
        backgroundColor: 'rgba(0,0,0,0.55)',
        display: 'flex', alignItems: 'flex-end',
        animation: 'fadeIn 0.3s ease',
      }}>
        <div style={{
          width: '100%',
          backgroundColor: 'white',
          borderRadius: '20px 20px 0 0',
          padding: '28px 24px 36px',
          boxShadow: '0 -8px 32px rgba(0,0,0,0.25)',
        }}>
          <div style={{ textAlign: 'center', marginBottom: '20px' }}>
            <div style={{ fontSize: '48px', marginBottom: '8px', animation: 'pulse 1.5s infinite' }}>⏰</div>
            <p style={{ fontWeight: 800, fontSize: '20px', color: '#e65100', margin: 0 }}>
              ¡Reporte pendiente!
            </p>
            <p style={{ fontSize: '14px', color: '#888', marginTop: '6px' }}>
              Han pasado <strong style={{ color: '#e65100' }}>
                {Math.floor(minutosSinReporte / 60)}h {minutosSinReporte % 60}min
              </strong> sin reporte de avance.
            </p>
            <p style={{ fontSize: '13px', color: '#aaa', marginTop: '4px' }}>
              El procedimiento indica reportar cada 2 horas.
            </p>
          </div>

          <button
            onClick={() => { setMinutosSinReporte(0); setModalParcial(true) }}
            style={{
              width: '100%', padding: '18px', borderRadius: '14px',
              backgroundColor: '#F8A12F', color: 'white',
              border: 'none', fontSize: '17px', fontWeight: 700,
              cursor: 'pointer', marginBottom: '12px',
            }}>
            📋 Registrar avance ahora
          </button>
        </div>
      </div>
    )}
    </>
  )
}
