import { useState, useEffect, useRef } from 'react'
import { useForm } from 'react-hook-form'
import { useMsal } from '../hooks/useMsal'
import { useApp } from '../context/AppContext'
import { Header } from '../components/Header'
import { updateListItem, getListItems, uploadAttachment } from '../services/sharepoint'
import { encolarOperacion } from '../services/indexedDB'
import { useProduccionParcial } from '../hooks/useProduccionParcial'
import { useMaestros } from '../hooks/useMaestros'
import { Toast } from '../components/Toast'

const inputStyle = {
  width: '100%', padding: '14px 12px', borderRadius: '10px',
  border: '2px solid #ddd', fontSize: '20px', fontWeight: 700,
  textAlign: 'right', backgroundColor: 'white',
  minHeight: '56px', fontVariantNumeric: 'tabular-nums', color: '#1a1a1a',
}

const smallInput = {
  width: '100%', padding: '10px 10px', borderRadius: '8px',
  border: '2px solid #ddd', fontSize: '15px', fontWeight: 600,
  textAlign: 'right', backgroundColor: 'white', color: '#1a1a1a',
}

export default function CierreTurno() {
  const { getToken, logout } = useMsal()
  const { turnoActivo, limpiarTurno, pendingCount, setPantalla, actualizarTurnoLocal } = useApp()
  const { productos: catalogoPT, materiasPrimas } = useMaestros(getToken)

  // Resolver código o nombre → nombre para display
  const resolverNombre = (valor) => {
    const catalogo = [...catalogoPT, ...materiasPrimas]
    const item = catalogo.find(p =>
      (p.Codigo || '') === valor ||
      (p.Nombre || p.Title || '').toLowerCase() === (valor || '').toLowerCase()
    )
    return item ? (item.Nombre || item.Title || valor) : valor
  }

  // Determina si un insumo es Base o Colorante — busca por código O nombre
  const getTipoInsumo = (insumoVal) => {
    const mp = materiasPrimas.find(p =>
      (p.Codigo || '') === insumoVal ||
      (p.Nombre || p.Title || '').toLowerCase() === (insumoVal || '').toLowerCase()
    )
    if (!mp) return 'Base'
    const tipo = (mp.TipoProducto || '').toUpperCase()
    return tipo === 'MC' ? 'Colorante' : 'Base'
  }
  const { register, handleSubmit, formState: { errors, isValid }, watch, setValue } = useForm({ mode: 'onChange' })
  const [enviando, setEnviando]       = useState(false)
  const [feedback, setFeedback]       = useState(null)
  const [toast, setToast]             = useState(null)
  const [fotoEvidencia, setFotoEvidencia] = useState(null)
  const fotoRef = useRef(null)
  // null = no respondido | false = cierre final | true = continúa en otra máquina
  const [transferenciaPendiente, setTransferenciaPendiente] = useState(null)

  // Kardex entries de esta OF — fuente de verdad de MP
  const [kardexOF, setKardexOF]   = useState([])
  const [mpEdits, setMpEdits]     = useState({}) // { [kardexID]: {KgUsado, KgMermaRec, KgMermaNoRec, KgDevueltos} }
  const [cargandoKardex, setCargandoKardex] = useState(false)
  // Acumulado real ya usado en registros cerrados anteriores de la misma OF
  const [mpYaUsadoOF, setMpYaUsadoOF] = useState(0)
  // Colorantes: colapsados por defecto, expandibles
  const [colorantesExpand, setColorantesExpand] = useState({})
  const [intentoEnviar, setIntentoEnviar] = useState(false)

  const veracidad      = watch('CheckboxVeracidad')
  const undConfWatch   = watch('UnidadesConformes')
  const undDefWatch    = watch('UnidadesDefectuosas')

  // Acumulado de reportes de avance previos (el delta lo ingresa el operario)
  const { acumConfTurno, acumDefTurno } = useProduccionParcial({
    getToken, registroId: turnoActivo?.spId, codigoLote: turnoActivo?.Codigo_Lote,
  })

  // Totales reales del turno = acumulado avances + delta de este cierre
  const totalConfTurnoFull = (acumConfTurno || 0) + (parseInt(undConfWatch) || 0)
  const totalDefTurnoFull  = (acumDefTurno  || 0) + (parseInt(undDefWatch)  || 0)
  const prodTotal     = totalConfTurnoFull + totalDefTurnoFull
  const pctDefWatch   = prodTotal > 0 ? (totalDefTurnoFull / prodTotal) * 100 : 0
  const fotoRequerida = pctDefWatch > 5
  const tieneReportes = acumConfTurno > 0 || acumDefTurno > 0

  // Cargar Kardex de la OF
  useEffect(() => {
    if (!turnoActivo?.Numero_OF) return
    let cancelled = false
    const cargar = async () => {
      setCargandoKardex(true)
      try {
        const token = await getToken()
        if (!token) return
        const [kardexTodos, regsTodos] = await Promise.all([
          getListItems(token, 'Kardex_MP', { top: 200 }),
          getListItems(token, 'Registro_Produccion', { top: 200 }),
        ])
        if (cancelled) return
        const delOF = kardexTodos.filter(k =>
          k.Numero_OF === turnoActivo.Numero_OF &&
          (!turnoActivo.Producto || !k.Producto || k.Producto === turnoActivo.Producto)
        )
        setKardexOF(delOF)
        // Sumar MP_KgUsado de registros cerrados anteriores (excluir el turno actual)
        const yaUsado = regsTodos
          .filter(r => r.Numero_OF === turnoActivo.Numero_OF &&
                       r.Estado === 'cerrado' &&
                       r.ID !== turnoActivo.spId)
          .reduce((s, r) => s + (r.MP_KgUsado || r.MP_Consumida_Calc || 0), 0)
        setMpYaUsadoOF(yaUsado)
        // Inicializar edits con valores existentes
        const edits = {}
        delOF.forEach(k => {
          edits[k.ID] = { KgUsado: '', KgMermaRec: '', KgMermaNoRec: '', KgDevueltos: '' }
        })
        setMpEdits(edits)
      } catch (err) {
        console.error('Error cargando Kardex:', err)
      } finally {
        if (!cancelled) setCargandoKardex(false)
      }
    }
    cargar()
    return () => { cancelled = true }
  }, [turnoActivo?.Numero_OF])

  // Guarda como string para permitir tipeo fluido de 0, 0.5, etc.
  // La conversión a número solo ocurre al calcular y al guardar en SP
  const updateMpEdit = (id, campo, valor) =>
    setMpEdits(prev => ({ ...prev, [id]: { ...prev[id], [campo]: valor } }))

  // Helper para leer valor numérico de un edit (string → float)
  const numEdit = (e, campo) => parseFloat(e?.[campo] ?? 0) || 0

  // Tipo determinado desde Maestro_Productos.TipoProducto (MP=Base, MC=Colorante)
  const mpBaseKardex = kardexOF.filter(k => getTipoInsumo(k.Insumo) === 'Base')
  // MP_KgUsado = lo que el operario declaró usar de MP Base (real, va a Registro_Produccion)
  const mpBaseTotal  = mpBaseKardex.reduce((s, k) => s + numEdit(mpEdits[k.ID], 'KgUsado'), 0)

  const hayBase = kardexOF.some(k => getTipoInsumo(k.Insumo) === 'Base')
  const mpColorantesKardex = kardexOF.filter(k => getTipoInsumo(k.Insumo) === 'Colorante')
  const mpBaseOk = mpBaseKardex.length === 0 || mpBaseKardex.every(k => numEdit(mpEdits[k.ID], 'KgUsado') > 0)
  const mpColorantesOk = mpColorantesKardex.length === 0 || mpColorantesKardex.some(k => numEdit(mpEdits[k.ID], 'KgUsado') > 0)

  // Ningún insumo puede tener saldo negativo — la suma de consumos no puede superar lo entregado
  const mpSaldosOk = kardexOF.length === 0 || (() => {
    const grupos = {}
    kardexOF.forEach(k => {
      const key = (k.Insumo || '').trim().toLowerCase()
      if (!grupos[key]) grupos[key] = []
      grupos[key].push(k)
    })
    return Object.values(grupos).every(grupo => {
      const k = grupo[0]
      const kgBase = grupo.reduce((s, entry) => s + (entry.KgDeclaradoOperario ?? entry.KgEntregados ?? 0), 0)
      const kgKardexAcum = grupo.reduce((s, entry) =>
        s + (entry.KgUsado || 0) + (entry.KgMermaRec || 0) + (entry.KgMermaNoRec || 0) + (entry.KgDevueltos || 0), 0)
      const esBase = getTipoInsumo(k.Insumo) === 'Base'
      const kgYaReg = kgKardexAcum > 0 ? kgKardexAcum : (esBase && mpYaUsadoOF > 0 ? mpYaUsadoOF : 0)
      const e = mpEdits[k.ID] || {}
      const balance = kgBase - kgYaReg
        - numEdit(e, 'KgUsado') - numEdit(e, 'KgMermaRec')
        - numEdit(e, 'KgMermaNoRec') - numEdit(e, 'KgDevueltos')
      return balance >= -0.01
    })
  })()

  const kgPorUnidad  = turnoActivo?.KgPorUnidadProducto ?? 0
  // Teórico = unidades producidas × estándar
  const mpTeorica   = kgPorUnidad > 0 ? prodTotal * kgPorUnidad : 0
  // Diferencia = |real - teórico| / real × 100
  const prodTeorica = kgPorUnidad > 0 && mpBaseTotal > 0 ? mpBaseTotal / kgPorUnidad : 0
  const diferencia  = mpBaseTotal > 0 && mpTeorica > 0
    ? Math.abs(mpBaseTotal - mpTeorica) / mpBaseTotal * 100 : 0

  // Al éxito: limpiar turno y volver a bienvenida (las opciones viven allí)
  useEffect(() => {
    if (feedback === 'exito') limpiarTurno()
  }, [feedback])

  const requiereRespuestaTransf = kardexOF.length > 0 && transferenciaPendiente === null
  // Bloquear campos MP hasta que el operario responda la pregunta de balance
  const mpDisabled = kardexOF.length > 0 && transferenciaPendiente === null
  const canSubmit = isValid && veracidad && !enviando && (!fotoRequerida || fotoEvidencia) && !requiereRespuestaTransf && mpSaldosOk

  const onSubmit = async (data) => {
    setIntentoEnviar(true)
    if (fotoRequerida && !fotoEvidencia) {
      setToast({ mensaje: `La foto es obligatoria cuando las defectuosas superan el 5% (${pctDefWatch.toFixed(1)}%).`, tipo: 'warn' })
      return
    }
    if (kardexOF.length > 0 && (!mpBaseOk || !mpColorantesOk || !mpSaldosOk)) return
    setEnviando(true)
    const ahora = new Date()

    // Total = acumulado de avances previos + delta ingresado en este cierre
    const undConformes   = (acumConfTurno || 0) + (parseInt(data.UnidadesConformes,   10) || 0)
    const undDefectuosas = (acumDefTurno  || 0) + (parseInt(data.UnidadesDefectuosas, 10) || 0)
    const undSueltas     = parseInt(data.UnidadesSueltas,     10) || 0
    const gruposConf     = parseInt(data.GruposConformes,     10) || 0
    const prodTotalFinal = undConformes + undDefectuosas

    // Semáforo usando Kardex base
    const minutosRetraso = turnoActivo.HoraInicio
      ? Math.round((ahora - new Date(turnoActivo.HoraInicio)) / 60000) : 0

    let estadoValidacion = 'Verde'
    let obsPcpAuto = ''

    if (undConformes === 0 && prodTotalFinal === 0) {
      estadoValidacion = 'Amarillo'
      obsPcpAuto = 'Auto: campos de producción vacíos'
    } else if (minutosRetraso > 240) {
      estadoValidacion = 'Amarillo'
      obsPcpAuto = `Auto: cierre con ${Math.round(minutosRetraso / 60)}h de retraso`
    } else if (diferencia > 8) {
      estadoValidacion = 'Rojo'
      obsPcpAuto = `Auto: diferencia MP vs producción ${diferencia.toFixed(1)}%`
    }

    // Saldo MP restante = total entregado − todo lo consumido en la OF (acumulado + este cierre)
    const mpBaseKardex = kardexOF.filter(k => getTipoInsumo(k.Insumo) === 'Base')

    const mpTotalEntregada = mpBaseKardex.reduce((s, k) => {
      // KgEntregados es la fuente oficial (PCP). Solo usar KgDeclaradoOperario si KgEntregados=0
      // (caso de entradas creadas por el operario sin registro previo de PCP)
      const kg = (k.KgEntregados > 0) ? k.KgEntregados : (k.KgDeclaradoOperario || 0)
      return s + kg
    }, 0)

    const mpTotalUsada = mpBaseKardex.reduce((s, k) => {
      const e = mpEdits[k.ID] || {}
      // KgUsado en Kardex ya acumula cierres anteriores de esta OF
      // + lo que el operario ingresa en este cierre
      return s
        + (k.KgUsado      || 0) + numEdit(e, 'KgUsado')
        + (k.KgMermaRec   || 0) + numEdit(e, 'KgMermaRec')
        + (k.KgMermaNoRec || 0) + numEdit(e, 'KgMermaNoRec')
        + (k.KgDevueltos  || 0) + numEdit(e, 'KgDevueltos')
    }, 0)

    const kgMPRestante = Math.max(0, mpTotalEntregada - mpTotalUsada)

    // ── Tiempos del turno ─────────────────────────────────────────────────────
    const horaInicio     = turnoActivo.HoraInicio ? new Date(turnoActivo.HoraInicio) : ahora
    const tiempoTurnoMin = parseFloat(((ahora - horaInicio) / 60000).toFixed(2))

    const paradasArr     = (() => { try { return JSON.parse(turnoActivo.Paradas || '[]') } catch { return [] } })()
    const tiempoParadaMin = parseFloat(paradasArr.reduce((s, p) =>
      s + (p.duracion_segundos != null ? p.duracion_segundos / 60 : (p.duracion_minutos || 0)), 0
    ).toFixed(2))
    const tiempoProductivoMin = parseFloat(Math.max(0, tiempoTurnoMin - tiempoParadaMin).toFixed(2))
    // ─────────────────────────────────────────────────────────────────────────

    // Nivel 1 — mínimo absoluto: solo Estado + HoraFin (100% garantizados)
    const p1 = {
      Estado:  'cerrado',
      HoraFin: ahora.toISOString(),
    }
    // Nivel 2 — datos de producción core + MP (columnas que deben existir en SP)
    const insumoBaseNombre = mpBaseKardex.filter(k => numEdit(mpEdits[k.ID], 'KgUsado') > 0)
      .map(k => k.Insumo).join(', ') || mpBaseKardex.map(k => k.Insumo).join(', ')

    const p2 = {
      CheckboxVeracidad:   true,
      UnidadesConformes:   undConformes,
      UnidadesDefectuosas: undDefectuosas,
      Estado_Validacion:   estadoValidacion,
      Paradas:             turnoActivo.Paradas || '[]',
      // MP — campos principales del registro de producción
      MP_KgUsado:          parseFloat(mpBaseTotal.toFixed(3)),
      Insumo_Base:         insumoBaseNombre,
    }
    // Nivel 3 — analíticos: pueden no existir todavía
    const p3 = {
      GruposConformes:       gruposConf,
      UnidadesSueltas:       undSueltas,
      Tiempo_Turno_Min:      tiempoTurnoMin,
      Tiempo_Parada_Min:     tiempoParadaMin,
      Tiempo_Productivo_Min: tiempoProductivoMin,
      MP_Consumida_Calc:     kgPorUnidad > 0 ? parseFloat((prodTotal * kgPorUnidad).toFixed(3)) : parseFloat(mpBaseTotal.toFixed(3)),
      Produccion_Teorica:    parseFloat(prodTeorica.toFixed(1)),
      Diferencia_Pct:        parseFloat(diferencia.toFixed(2)),
      Fecha_Validacion:      ahora.toISOString(),
      KgMPRestante:          parseFloat(kgMPRestante.toFixed(3)),
      ...(obsPcpAuto ? { Obs_PCP: obsPcpAuto } : {}),
    }

    const todoPayload = { ...p1, ...p2, ...p3 }

    try {
      const token = await getToken()
      if (turnoActivo.spId && navigator.onLine && token) {
        // Foto de defectuosas
        if (fotoEvidencia) {
          try {
            const nombreFoto = `defectuosos_${turnoActivo.spId}_${Date.now()}.jpg`
            await uploadAttachment(token, 'Registro_Produccion', turnoActivo.spId, nombreFoto, fotoEvidencia)
            p1.TieneFoto = true
          } catch { /* foto no crítica */ }
        }

        // Nivel 1 — DEBE guardarse (Estado + HoraFin)
        await updateListItem(token, 'Registro_Produccion', turnoActivo.spId, p1)

        // Nivel 2 — datos de producción
        try {
          await updateListItem(token, 'Registro_Produccion', turnoActivo.spId, p2)
        } catch (e2) {
          console.warn('CierreTurno p2 fallback:', e2.message)
          // Intentar campo a campo los más importantes
          for (const [k, v] of Object.entries(p2)) {
            try { await updateListItem(token, 'Registro_Produccion', turnoActivo.spId, { [k]: v }) }
            catch { /* campo no existe */ }
          }
        }

        // Nivel 3 — analíticos (campo a campo si el bloque falla)
        try {
          await updateListItem(token, 'Registro_Produccion', turnoActivo.spId, p3)
        } catch (e3) {
          console.warn('CierreTurno p3 fallback campo a campo:', e3.message)
          for (const [k, v] of Object.entries(p3)) {
            try { await updateListItem(token, 'Registro_Produccion', turnoActivo.spId, { [k]: v }) }
            catch { /* campo analítico no existe en SP */ }
          }
        }

        // Kardex MP
        const grupos = {}
        kardexOF.forEach(k => {
          const key = (k.Insumo || '').trim().toLowerCase()
          if (!grupos[key]) grupos[key] = []
          grupos[key].push(k)
        })
        try {
          await Promise.allSettled(
            Object.values(grupos).map(grupo => {
              const principal = grupo[0]
              const e = mpEdits[principal.ID]
              if (!e) return Promise.resolve()
              return updateListItem(token, 'Kardex_MP', principal.ID, {
                KgUsado:      (principal.KgUsado      || 0) + numEdit(e, 'KgUsado'),
                KgMermaRec:   (principal.KgMermaRec   || 0) + numEdit(e, 'KgMermaRec'),
                KgMermaNoRec: (principal.KgMermaNoRec || 0) + numEdit(e, 'KgMermaNoRec'),
                KgDevueltos:  (principal.KgDevueltos  || 0) + numEdit(e, 'KgDevueltos'),
              })
            })
          )
        } catch { /* Kardex no crítico */ }

      } else {
        await encolarOperacion({ tipo: 'update', listName: 'Registro_Produccion', spId: turnoActivo.spId, data: todoPayload })
      }
      setFeedback('exito')
    } catch (err) {
      console.error('CierreTurno error crítico:', err)
      await encolarOperacion({ tipo: 'update', listName: 'Registro_Produccion', spId: turnoActivo.spId, data: todoPayload })
      setFeedback('exito')
    } finally {
      setEnviando(false)
    }
  }

  return (
    <div style={{ backgroundColor: '#f0f2f5', minHeight: '100vh' }}>
      <Toast toast={toast} onClose={() => setToast(null)} />
      <Header
        titulo="Registro de producción"
        subtitulo={`${turnoActivo?.Operario} · ${resolverNombre(turnoActivo?.Producto)}${turnoActivo?.Color ? ' · ' + turnoActivo.Color : ''}`}
        pendingCount={pendingCount}
        onLogout={logout}
        color="#004895"
      />

      <div style={{ padding: '12px', maxWidth: '520px', margin: '0 auto', display: 'flex', flexDirection: 'column', gap: '10px' }}>

        <form onSubmit={handleSubmit(onSubmit)} style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>

          {/* ── Producción ── */}
          <div style={{ backgroundColor: 'white', borderRadius: '14px', overflow: 'hidden', boxShadow: '0 2px 8px rgba(0,0,0,0.08)' }}>
            {/* Header simple */}
            <div style={{ backgroundColor: '#004895', padding: '10px 14px' }}>
              <p style={{ fontSize: '10px', color: 'rgba(255,255,255,0.7)', textTransform: 'uppercase', letterSpacing: '0.06em', margin: 0 }}>1 — Producción del turno</p>
            </div>
            <div style={{ padding: '12px' }}>

            {/* Grid 2 columnas — misma estructura que balance de MP */}
            {(() => {
              const totalConfTurno = acumConfTurno + (parseInt(undConfWatch) || 0)
              const totalDef = acumDefTurno + (parseInt(undDefWatch) || 0)
              const campos = [
                { name: 'UnidadesConformes',   label: 'Und. conformes',   color: '#2e7d32', requerido: true,  total: totalConfTurno,
                  hint: tieneReportes ? `${acumConfTurno} (avances) + ${parseInt(undConfWatch)||0} = ${totalConfTurno}` : null },
                { name: 'UnidadesDefectuosas', label: 'Und. defectuosas',  color: '#c62828', requerido: true,  total: totalDef,
                  hint: tieneReportes && acumDefTurno > 0 ? `${acumDefTurno} (avances) + ${parseInt(undDefWatch)||0} = ${totalDef}` : null },
                { name: 'GruposConformes',     label: 'Grupos conformes',  color: '#555555', requerido: false, total: null,
                  hint: totalConfTurno > 0 ? `De las ${totalConfTurno} und. conformes del turno` : null },
                { name: 'UnidadesSueltas',     label: 'Und. sueltas',      color: '#555555', requerido: true,  total: null,
                  hint: totalConfTurno > 0 ? `Conformes sin grupo (${totalConfTurno} und. totales)` : null },
              ]
              return (
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px' }}>
                  {campos.map(({ name, label, color, requerido: req, hint, total }) => (
                    <div key={name} style={{ backgroundColor: '#fafafa', borderRadius: '10px', border: `1.5px solid ${color}30`, padding: '10px' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
                        <label style={{ fontSize: '11px', fontWeight: 700, color }}>{label}{req ? ' *' : ''}</label>
                        {total != null && total > 0 && (
                          <span style={{ fontSize: '13px', fontWeight: 800, color, backgroundColor: `${color}15`, borderRadius: '6px', padding: '1px 7px' }}>{total}</span>
                        )}
                      </div>
                      <input
                        {...register(name, {
                          required: req ? `${label} es obligatorio` : false,
                          min: { value: 0, message: 'No puede ser negativo' },
                          validate: v => v !== '' && !Number.isInteger(Number(v)) ? 'Debe ser entero' : true,
                        })}
                        type="number" inputMode="numeric" step="1" min="0" placeholder="0"
                        style={{ width: '100%', padding: '8px 10px', borderRadius: '8px', border: `2px solid ${errors[name] ? '#d32f2f' : color}50`, fontSize: '22px', fontWeight: 800, textAlign: 'right', color: '#1a1a1a', backgroundColor: 'white', boxSizing: 'border-box' }}
                      />
                      {hint && <p style={{ fontSize: '10px', color: '#888', margin: '4px 0 0', lineHeight: 1.3 }}>{hint}</p>}
                      {errors[name] && <p style={{ color: '#d32f2f', fontSize: '10px', margin: '3px 0 0' }}>{errors[name].message}</p>}
                    </div>
                  ))}
                </div>
              )
            })()}

            {/* Foto si defectuosas > 5% */}
            {fotoRequerida && (
              <div>
                <label style={{ fontWeight: 600, fontSize: '14px', color: '#c62828', display: 'block', marginBottom: '6px' }}>
                  Foto obligatoria (defectuosas {pctDefWatch.toFixed(1)}% &gt; 5%) *
                </label>
                <button type="button" onClick={() => fotoRef.current?.click()} style={{
                  width: '100%', padding: '14px', borderRadius: '10px',
                  border: `2px dashed ${fotoEvidencia ? '#4caf50' : '#c62828'}`,
                  backgroundColor: fotoEvidencia ? '#e8f5e9' : 'white',
                  color: fotoEvidencia ? '#2e7d32' : '#c62828', fontSize: '14px',
                  fontWeight: 600, cursor: 'pointer', minHeight: '52px',
                }}>
                  {fotoEvidencia ? `✓ ${fotoEvidencia.name}` : '📷 Adjuntar foto de defectuosas'}
                </button>
                <input ref={fotoRef} type="file" accept="image/*" capture="environment"
                  style={{ display: 'none' }}
                  onChange={e => setFotoEvidencia(e.target.files?.[0] || null)} />
              </div>
            )}
            </div>{/* fin padding producción */}
          </div>

          {/* ── Balance de MP ── */}
          <div style={{ backgroundColor: 'white', borderRadius: '14px', overflow: 'hidden', boxShadow: '0 2px 8px rgba(0,0,0,0.08)' }}>
            <div style={{ backgroundColor: '#1b5e20', padding: '10px 14px' }}>
              <p style={{ fontSize: '10px', color: 'rgba(255,255,255,0.7)', textTransform: 'uppercase', letterSpacing: '0.06em', margin: 0 }}>2 — Balance de materia prima</p>
            </div>
            <div style={{ padding: '12px' }}>

            {/* Pregunta: ¿la OF continúa en otra máquina? */}
            {kardexOF.length > 0 && (
              <div style={{
                borderRadius: '10px', marginBottom: '14px', overflow: 'hidden',
                border: `2px solid ${transferenciaPendiente === null ? '#F8A12F' : transferenciaPendiente ? '#37BEEC' : '#4caf50'}`,
              }}>
                <div style={{
                  backgroundColor: transferenciaPendiente === null ? '#fff8e1' : transferenciaPendiente ? '#e3f7fd' : '#e8f5e9',
                  padding: '10px 14px',
                }}>
                  <p style={{ fontSize: '13px', fontWeight: 700, color: '#333', margin: '0 0 10px' }}>
                    ¿La producción de esta OF continúa en otra máquina?
                  </p>
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <button type="button" onClick={() => setTransferenciaPendiente(false)} style={{
                      flex: 1, padding: '10px 8px', borderRadius: '8px', border: '2px solid',
                      borderColor: transferenciaPendiente === false ? '#4caf50' : '#ddd',
                      backgroundColor: transferenciaPendiente === false ? '#4caf50' : 'white',
                      color: transferenciaPendiente === false ? 'white' : '#333',
                      fontSize: '13px', fontWeight: 700, cursor: 'pointer',
                    }}>
                      ✓ No — cierre color/turno
                    </button>
                    <button type="button" onClick={() => setTransferenciaPendiente(true)} style={{
                      flex: 1, padding: '10px 8px', borderRadius: '8px', border: '2px solid',
                      borderColor: transferenciaPendiente === true ? '#37BEEC' : '#ddd',
                      backgroundColor: transferenciaPendiente === true ? '#37BEEC' : 'white',
                      color: transferenciaPendiente === true ? 'white' : '#333',
                      fontSize: '13px', fontWeight: 700, cursor: 'pointer',
                    }}>
                      → Sí — se transfiere
                    </button>
                  </div>
                  {transferenciaPendiente === true && (
                    <p style={{ fontSize: '11px', color: '#0288d1', marginTop: '8px', marginBottom: 0 }}>
                      ℹ No llenes "Kg devueltos" — PCP gestionará los kg restantes desde Kardex MP.
                    </p>
                  )}
                </div>
              </div>
            )}

            {cargandoKardex ? (
              <p style={{ color: '#888', fontSize: '13px' }}>Cargando insumos de la OF...</p>
            ) : kardexOF.length === 0 ? (
              <div style={{ backgroundColor: '#fff8e1', borderRadius: '8px', padding: '12px', fontSize: '13px', color: '#e65100' }}>
                ⚠ No hay MP registrada en Kardex para esta OF ({turnoActivo?.Numero_OF}).
                PCP debe registrar la entrega en Kardex antes o después.
              </div>
            ) : (
              // Consolidar entradas del mismo insumo en una tarjeta
              (() => {
                const grupos = {}
                kardexOF.forEach(k => {
                  const key = (k.Insumo || '').trim().toLowerCase()
                  if (!grupos[key]) grupos[key] = []
                  grupos[key].push(k)
                })
                return Object.values(grupos)
                  .sort((ga, gb) => {
                    const ta = getTipoInsumo(ga[0].Insumo) === 'Base' ? 0 : 1
                    const tb = getTipoInsumo(gb[0].Insumo) === 'Base' ? 0 : 1
                    return ta - tb
                  })
                  .map(grupo => {
                    // Representante principal (primer item, generalmente el de PCP)
                    const k = grupo[0]
                    // Total entregado por PCP para este insumo
                    const kgBase = grupo.reduce((s, entry) => s + (entry.KgDeclaradoOperario ?? entry.KgEntregados ?? 0), 0)
                    // Ya usado en Kardex (registros anteriores de la OF)
                    const kgKardexAcum = grupo.reduce((s, entry) =>
                      s + (entry.KgUsado || 0) + (entry.KgMermaRec || 0) + (entry.KgMermaNoRec || 0) + (entry.KgDevueltos || 0), 0)
                    // Si Kardex no tiene acumulado actualizado, usar Registro_Produccion como fallback
                    const esBase = getTipoInsumo(k.Insumo) === 'Base'
                    const kgYaReg = kgKardexAcum > 0 ? kgKardexAcum
                      : (esBase && mpYaUsadoOF > 0 ? mpYaUsadoOF : 0)
                    // Edits: usar el ID del primer item como clave de edición
                    const e = mpEdits[k.ID] || {}
                    const balance = kgBase
                      - kgYaReg
                      - numEdit(e, 'KgUsado') - numEdit(e, 'KgMermaRec')
                      - numEdit(e, 'KgMermaNoRec') - numEdit(e, 'KgDevueltos')
                const ok = Math.abs(balance) < 0.01
                const tipo = getTipoInsumo(k.Insumo)
                return (
                  <div key={k.ID} style={{
                    marginBottom: '14px', borderRadius: '10px', overflow: 'hidden',
                    border: `1.5px solid ${tipo === 'Base' ? '#c5cae9' : '#c8e6c9'}`,
                  }}>
                    {/* Encabezado MP */}
                    <div
                      onClick={tipo !== 'Base' ? () => setColorantesExpand(prev => ({ ...prev, [k.ID]: !prev[k.ID] })) : undefined}
                      style={{
                        backgroundColor: tipo === 'Base' ? '#3949ab' : '#2e7d32',
                        color: 'white', padding: '8px 12px',
                        display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start',
                        cursor: tipo !== 'Base' ? 'pointer' : 'default',
                      }}>
                      <div>
                        <div style={{ fontWeight: 700, fontSize: '14px' }}>
                          {tipo !== 'Base' && <span style={{ marginRight: '6px', fontSize: '12px' }}>{colorantesExpand[k.ID] ? '▼' : '▶'}</span>}
                          {resolverNombre(k.Insumo)}
                        </div>
                        <div style={{ fontSize: '11px', opacity: 0.8, marginTop: '2px', display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                          <span>PCP: {k.KgEntregados} kg</span>
                          {k.KgDeclaradoOperario != null && Math.abs((k.KgDeclaradoOperario || 0) - (k.KgEntregados || 0)) >= 0.01 && (
                            <span style={{ backgroundColor: 'rgba(255,200,0,0.35)', borderRadius: '4px', padding: '0 5px' }}>
                              Op: {k.KgDeclaradoOperario} kg
                            </span>
                          )}
                        </div>
                        {kgYaReg > 0 && (
                          <div style={{ fontSize: '11px', opacity: 0.75, marginTop: '2px' }}>
                            Usado prev.: {kgYaReg.toFixed(2)} kg
                          </div>
                        )}
                      </div>
                      <span style={{
                        fontSize: '12px', fontWeight: 700, flexShrink: 0, marginLeft: '8px',
                        backgroundColor: ok ? 'rgba(255,255,255,0.2)' : 'rgba(255,100,100,0.4)',
                        borderRadius: '6px', padding: '3px 8px', textAlign: 'right',
                      }}>
                        Saldo: {balance.toFixed(2)} kg
                      </span>
                    </div>

                    {/* Campos de balance — Base siempre visible, Colorante solo si expandido */}
                    {(tipo === 'Base' || colorantesExpand[k.ID]) && (
                    <div style={{ backgroundColor: 'white', padding: '12px', position: 'relative' }}>
                      {mpDisabled && (
                        <div style={{
                          position: 'absolute', inset: 0, zIndex: 2, borderRadius: '0 0 8px 8px',
                          backgroundColor: 'rgba(245,245,245,0.82)',
                          display: 'flex', alignItems: 'center', justifyContent: 'center',
                        }}>
                          <span style={{ fontSize: '12px', color: '#888', fontWeight: 600 }}>
                            ⬆ Responde la pregunta de arriba
                          </span>
                        </div>
                      )}
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px' }}>
                      {[
                        ['KgUsado',      'Kg usado en prod.'],
                        ['KgMermaRec',   'Kg merma rec.'],
                        ['KgMermaNoRec', 'Kg merma no rec.'],
                        ['KgDevueltos',  'Kg devueltos'],
                      ].filter(([campo]) => !(campo === 'KgDevueltos' && transferenciaPendiente === true))
                      .map(([campo, label]) => {
                        const esBaseReq = campo === 'KgUsado' && tipo === 'Base'
                        const campoError = esBaseReq && intentoEnviar && numEdit(e, 'KgUsado') <= 0
                        return (
                        <div key={campo}>
                          <label style={{ fontSize: '11px', fontWeight: 600, color: campoError ? '#d32f2f' : '#555', display: 'block', marginBottom: '3px' }}>
                            {label}{esBaseReq ? ' *' : ''}
                          </label>
                          <input
                            type="number" step="0.01" min="0"
                            value={e[campo] ?? ''}
                            onChange={ev => updateMpEdit(k.ID, campo, ev.target.value)}
                            placeholder="0"
                            disabled={mpDisabled}
                            style={{ ...smallInput, borderColor: campoError ? '#d32f2f' : '#ddd', opacity: mpDisabled ? 0.5 : 1 }}
                          />
                          {campoError && <p style={{ color: '#d32f2f', fontSize: '10px', margin: '2px 0 0' }}>Obligatorio</p>}
                        </div>
                        )
                      })}
                      </div>
                      {balance < -0.01 && (
                        <div style={{ marginTop: '10px', backgroundColor: '#ffebee', borderRadius: '8px', padding: '8px 12px', fontSize: '12px', color: '#c62828', fontWeight: 700 }}>
                          ⚠ Excede el saldo disponible. Máximo: <strong>{(kgBase - kgYaReg).toFixed(2)} kg</strong> en total (usado + mermas + devueltos).
                        </div>
                      )}
                    </div>
                    )}
                  </div>
                )
              })
              })()
            )}

            {intentoEnviar && !mpColorantesOk && mpColorantesKardex.length > 0 && (
              <div style={{ backgroundColor: '#ffebee', borderRadius: '8px', padding: '8px 12px', fontSize: '12px', color: '#c62828', margin: '4px 0 8px' }}>
                ⚠ Expande al menos un colorante e ingresa los kg usados
              </div>
            )}

            {/* Aviso si no hay MP Base etiquetada */}
            {kardexOF.length > 0 && !hayBase && (
              <div style={{ backgroundColor: '#fff8e1', borderRadius: '8px', padding: '10px 12px', fontSize: '12px', color: '#e65100', marginTop: '8px' }}>
                ⚠ Ningún insumo tiene Tipo=<strong>Base</strong>. El semáforo no puede calcularse.
                Asegúrate de que los insumos estructurales (PP Natural, PP Clarificado, etc.)
                tengan <strong>Tipo = Base</strong> en Kardex.
              </div>
            )}

            {/* Semáforo calculado */}
            {kardexOF.length > 0 && hayBase && (
              <div style={{
                marginTop: '8px', padding: '10px 14px', borderRadius: '8px',
                backgroundColor: diferencia > 8 ? '#ffebee' : '#e8f5e9',
                border: `1px solid ${diferencia > 8 ? '#f44336' : '#4caf50'}`,
                fontSize: '13px',
              }}>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span style={{ color: '#555' }}>MP real (declarada):</span>
                  <strong style={{ color: '#1a1a1a' }}>{mpBaseTotal.toFixed(2)} kg</strong>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span style={{ color: '#555' }}>MP teórica ({prodTotal} und × {kgPorUnidad} kg):</span>
                  <strong style={{ color: '#1a1a1a' }}>{mpTeorica > 0 ? `${mpTeorica.toFixed(2)} kg` : '—'}</strong>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: 700, borderTop: '1px solid #ddd', paddingTop: '6px', marginTop: '4px' }}>
                  <span style={{ color: diferencia > 8 ? '#c62828' : '#2e7d32' }}>Diferencia:</span>
                  <strong style={{ color: diferencia > 8 ? '#c62828' : '#2e7d32' }}>
                    {diferencia.toFixed(1)}% {diferencia > 8 ? '⚠ ROJO' : '✓ OK'}
                  </strong>
                </div>
              </div>
            )}
            </div>{/* fin padding MP */}
          </div>

          {/* Error MP visible antes del botón */}
          {intentoEnviar && kardexOF.length > 0 && (!mpBaseOk || !mpColorantesOk || !mpSaldosOk) && (
            <div style={{ backgroundColor: '#ffebee', border: '2px solid #d32f2f', borderRadius: '10px', padding: '12px 14px', fontSize: '13px', color: '#c62828', fontWeight: 600 }}>
              ⚠ Completa el balance de MP antes de confirmar:
              {!mpBaseOk && <div style={{ fontWeight: 400, marginTop: '4px' }}>• Ingresa los kg usados en cada MP Base</div>}
              {!mpColorantesOk && mpColorantesKardex.length > 0 && <div style={{ fontWeight: 400, marginTop: '4px' }}>• Ingresa los kg usados en al menos un colorante</div>}
              {!mpSaldosOk && <div style={{ fontWeight: 400, marginTop: '4px' }}>• Uno o más insumos superan el saldo entregado por PCP</div>}
            </div>
          )}

          {/* ── 3: Confirmación + envío ── */}
          <div style={{ backgroundColor: 'white', borderRadius: '14px', overflow: 'hidden', boxShadow: '0 2px 8px rgba(0,0,0,0.08)' }}>
            <div style={{ backgroundColor: canSubmit ? '#2e7d32' : '#888', padding: '10px 14px' }}>
              <p style={{ fontSize: '10px', color: 'rgba(255,255,255,0.7)', textTransform: 'uppercase', letterSpacing: '0.06em', margin: 0 }}>3 — Confirmar y enviar</p>
            </div>
            <div style={{ padding: '14px' }}>
              <label style={{ display: 'flex', alignItems: 'flex-start', gap: '12px', cursor: 'pointer', marginBottom: '12px' }}>
                <input {...register('CheckboxVeracidad', { required: true })}
                  type="checkbox"
                  style={{ width: '22px', height: '22px', marginTop: '2px', accentColor: '#004895', flexShrink: 0 }}
                />
                <span style={{ fontSize: '14px', color: '#333', lineHeight: 1.4 }}>
                  Declaro que los datos son correctos y corresponden a la producción real del turno.
                </span>
              </label>

              <button type="submit" disabled={!canSubmit} style={{
                backgroundColor: canSubmit ? '#2e7d32' : '#bbb',
                color: 'white', border: 'none', borderRadius: '12px',
                padding: '16px', fontSize: '17px', fontWeight: 800,
                minHeight: '58px', width: '100%',
                cursor: canSubmit ? 'pointer' : 'not-allowed',
                boxShadow: canSubmit ? '0 4px 12px rgba(46,125,50,0.3)' : 'none',
              }}>
                {enviando ? '⏳ Guardando...' : '✓ Confirmar registro de producción'}
              </button>
            </div>
          </div>

          <button type="button" onClick={() => setPantalla('turno-activo')} style={{
            backgroundColor: 'transparent', color: '#888',
            border: '1.5px solid #ddd', borderRadius: '12px',
            padding: '12px', fontSize: '14px', cursor: 'pointer',
          }}>
            ← Volver al turno
          </button>

        </form>
      </div>
    </div>
  )
}
