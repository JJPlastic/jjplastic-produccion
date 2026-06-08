import { useState, useEffect, useCallback } from 'react'
import { format, parseISO, differenceInHours } from 'date-fns'
import { es } from 'date-fns/locale'
import { useMsal } from '../../hooks/useMsal'
import { getListItems, updateListItem, resolveUserId } from '../../services/sharepoint'

// ─── Clave de turno para agrupar registros del mismo turno ────────────────────
// eslint-disable-next-line react-refresh/only-export-components
export const turnoKey = (r) => {
  const fecha = (r.Fecha || r.HoraInicio || '').split('T')[0]
  // Incluir Numero_OF para separar grupos de distintas OFs en el mismo turno
  return `${r.Title || r.Maquina || ''}|${r.Numero_OF || ''}|${fecha}|${r.Turno || ''}`
}

// ─── Fórmula semáforo agregada por turno ─────────────────────────────────────
// records: todos los registros del mismo Maquina+Fecha+Turno
// kgPorUnidadMap: { codigoProducto: kgPorUnidad }
// kgKardex: kg netos del Kardex para ese turno (puede ser null)
const calcularSemaforoTurno = (records, kgPorUnidadMap, kgKardex) => {
  const cerrados = records.filter(r => r.Estado === 'cerrado')
  if (cerrados.length === 0) return { estado: 'Amarillo', motivo: 'Turno abierto' }

  // Solo verificar campos de producción (los de MP viven en Kardex, no en Registro)
  const algoCamposVacios = cerrados.some(r => r.UnidadesConformes == null)
  if (algoCamposVacios) return { estado: 'Amarillo', motivo: 'Campos de producción vacíos' }

  // Gap 2: Foto requerida si defectuosas >5% y no se adjuntó
  const totalProd = cerrados.reduce((s, r) => s + (r.UnidadesConformes || 0) + (r.UnidadesDefectuosas || 0), 0)
  const totalDef  = cerrados.reduce((s, r) => s + (r.UnidadesDefectuosas || 0), 0)
  const pctDef = totalProd > 0 ? totalDef / totalProd * 100 : 0
  const fotoRequerida = pctDef > 5
  const tieneFoto = cerrados.some(r => r.TieneFoto === true)
  if (fotoRequerida && !tieneFoto) return { estado: 'Amarillo', motivo: `Foto requerida (${pctDef.toFixed(1)}% defectuosas sin foto)` }

  // Retraso: el último cierre fue hace más de 4h
  const ultimaFin = cerrados.map(r => r.HoraFin ? parseISO(r.HoraFin) : null).filter(Boolean)
  const masReciente = ultimaFin.length ? new Date(Math.max(...ultimaFin)) : null
  if (masReciente && differenceInHours(new Date(), masReciente) > 4) {
    return { estado: 'Amarillo', motivo: `Retraso ${Math.round(differenceInHours(new Date(), masReciente))}h` }
  }

  // Agregados del turno (solo campos de producción — MP está en Kardex)
  const totConf = cerrados.reduce((s, r) => s + (r.UnidadesConformes  || 0), 0)
  const totDef  = cerrados.reduce((s, r) => s + (r.UnidadesDefectuosas|| 0), 0)

  // MP real = MP_KgUsado declarado por operario (Base MP)
  // MP teórica = MP_Consumida_Calc = (Conf+Def) × KgPorUnidad (calculado al cierre)
  const tieneKgUsado = cerrados.some(r => r.MP_KgUsado != null && r.MP_KgUsado > 0)
  const mpConsumida  = tieneKgUsado
    ? cerrados.reduce((s, r) => s + (r.MP_KgUsado || 0), 0)
    : (kgKardex ?? 0)

  // Para múltiples productos, promediar KgPorUnidad ponderado por producción
  // Si solo hay un producto: usar directo
  // Fórmula correcta para múltiples productos:
  // MP_teorica = Σ (UnidadesDeclaradas_i × KgPorUnidad_i)
  // Diferencia = |MP_consumida_real - MP_teorica| / MP_consumida_real × 100
  const mpTeorica = cerrados.reduce((s, r) => {
    const kpu = kgPorUnidadMap[r.Producto] || 0
    return s + ((r.UnidadesConformes || 0) + (r.UnidadesDefectuosas || 0)) * kpu
  }, 0)

  const tieneEstandar = cerrados.some(r => (kgPorUnidadMap[r.Producto] || 0) > 0)
  if (!tieneEstandar) return { estado: 'Verde', motivo: 'Sin estándar MP', mpConsumida, totConf, totDef }

  const diferenciaPct = mpConsumida > 0 ? Math.abs(mpConsumida - mpTeorica) / mpConsumida * 100 : 0

  const base = { mpConsumida, mpTeorica, diferenciaPct, kgKardex, totConf, totDef }
  if (diferenciaPct > 8) return { estado: 'Rojo', motivo: `Diferencia MP: ${diferenciaPct.toFixed(1)}%`, ...base }
  return { estado: 'Verde', motivo: `Diferencia MP: ${diferenciaPct.toFixed(1)}%`, ...base }
}

// Para compatibilidad con la lista (semáforo individual rápido)
const calcularSemaforo = (reg, kgPorUnidad) =>
  calcularSemaforoTurno([reg], { [reg.Producto]: kgPorUnidad }, null)

const COLORES_SEMAFORO = { Verde: '#2e7d32', Amarillo: '#f57f17', Rojo: '#c62828' }
const BG_SEMAFORO = { Verde: '#e8f5e9', Amarillo: '#fffde7', Rojo: '#ffebee' }

const Badge = ({ estado }) => (
  <span style={{
    backgroundColor: BG_SEMAFORO[estado] || '#f5f5f5',
    color: COLORES_SEMAFORO[estado] || '#555',
    border: `1.5px solid ${COLORES_SEMAFORO[estado] || '#ccc'}`,
    borderRadius: '20px', padding: '3px 12px',
    fontSize: '13px', fontWeight: 700,
  }}>
    {estado === 'Verde' ? '● Verde' : estado === 'Amarillo' ? '● Amarillo' : '● Rojo'}
  </span>
)

const LABEL = { fontWeight: 600, fontSize: '14px', display: 'block', marginBottom: '6px', color: '#222' }
const INPUT_STYLE = { width: '100%', padding: '10px', borderRadius: '8px', border: '2px solid #ddd', fontSize: '14px', fontFamily: 'inherit', color: '#1a1a1a', backgroundColor: 'white' }

// ─── Modal de corrección ──────────────────────────────────────────────────────
const ModalCorreccion = ({ registro, kgPorUnidad, kgKardex, itemsKardexTurno = [], registrosTurno, kgPorUnidadMap, tiposInsumo = {}, onGuardar, onCerrar, guardando }) => {
  // Filtrar solo insumos tipo Base (MP estructural — no colorantes)
  const itemsBase = itemsKardexTurno.filter(k => {
    const nombre = (k.Insumo || '').toLowerCase()
    const tipo = tiposInsumo[nombre] || tiposInsumo[nombre.split(' ')[0]] || 'Base'
    return tipo !== 'Colorante'
  })
  const [estado, setEstado] = useState(registro.Estado_Validacion || 'Verde')
  const obsInicial = (registro.Obs_PCP || '').startsWith('Auto:') ? '' : (registro.Obs_PCP || '')
  const [obs, setObs] = useState(obsInicial)
  const [motivo, setMotivo] = useState(registro.Motivo_Correccion || '')
  const [editando, setEditando] = useState(true)
  const [autorizadoPorJefe, setAutorizadoPorJefe] = useState(false)

  // Solo MP_KgUsado (real Base MP) editable en Registro_Produccion
  // Merma, devueltos → solo en Kardex_MP
  const [mpLineaEdits, setMpLineaEdits] = useState({
    MP_KgUsado: registro.MP_KgUsado != null ? String(registro.MP_KgUsado) : '',
  })

  // kardexEdits solo para compatibilidad con guardado en Kardex_MP
  const [kardexEdits, setKardexEdits] = useState(() => {
    const m = {}
    itemsBase.forEach(k => {
      m[k.ID] = {
        KgUsado:      k.KgUsado      != null ? String(k.KgUsado)      : '',
        KgMermaRec:   k.KgMermaRec   != null ? String(k.KgMermaRec)   : '',
        KgMermaNoRec: k.KgMermaNoRec != null ? String(k.KgMermaNoRec) : '',
        KgDevueltos:  k.KgDevueltos  != null ? String(k.KgDevueltos)  : '',
      }
    })
    return m
  })
  const updateKardex = (id, campo, valor) =>
    setKardexEdits(prev => ({ ...prev, [id]: { ...prev[id], [campo]: valor } }))

  // Detectar si es corrección retroactiva (fecha anterior a hoy)
  const fechaReg = registro.Fecha ? registro.Fecha.split('T')[0] : ''
  const esRetroactivo = fechaReg && fechaReg < new Date().toISOString().split('T')[0]
  // Solo campos que existen en Registro_Produccion (MP se gestiona en Kardex)
  const [edits, setEdits] = useState({
    UnidadesConformes:   registro.UnidadesConformes   ?? '',
    GruposConformes:     registro.GruposConformes     ?? '',
    UnidadesSueltas:     registro.UnidadesSueltas     ?? '',
    UnidadesDefectuosas: registro.UnidadesDefectuosas ?? '',
  })

  // Cálculo del turno completo para el análisis
  const turno = registrosTurno?.length > 1 ? registrosTurno : [registro]
  const analisis = calcularSemaforoTurno(turno, kgPorUnidadMap || { [registro.Producto]: kgPorUnidad }, kgKardex)
  const colorDif = analisis.diferenciaPct == null ? '#555'
    : analisis.diferenciaPct > 8 ? '#c62828' : '#2e7d32'
  const esMultiProducto = (registrosTurno?.length || 1) > 1

  return (
    <div style={{
      position: 'fixed', inset: 0, backgroundColor: 'rgba(0,0,0,0.65)',
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      zIndex: 300, padding: '12px',
    }}>
      <div style={{
        backgroundColor: 'white', borderRadius: '16px',
        width: '100%', maxWidth: '500px', maxHeight: '94vh',
        display: 'flex', flexDirection: 'column', overflow: 'hidden',
        boxShadow: '0 20px 60px rgba(0,0,0,0.3)',
      }}>
        {/* ── Header compacto ── */}
        <div style={{ backgroundColor: '#004895', padding: '12px 16px', color: 'white', flexShrink: 0 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <p style={{ fontWeight: 800, fontSize: '16px', margin: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                {registro.NombreProducto || registro.Producto || '—'}
              </p>
              {registro.Color && (
                <p style={{ fontSize: '12px', opacity: 0.85, margin: '1px 0 0', fontWeight: 600 }}>{registro.Color}</p>
              )}
              <p style={{ fontSize: '11px', opacity: 0.7, margin: '2px 0 0' }}>
                {registro.Title} · Turno {registro.Turno === 'M' ? 'Mañana' : registro.Turno === 'T' ? 'Tarde' : 'Noche'} · 👤 {registro.Operario}
                {registro.HoraInicio && <> · {registro.HoraInicio.split('T')[1]?.slice(0,5)}</>}
                {registro.HoraFin && <> → {registro.HoraFin.split('T')[1]?.slice(0,5)}</>}
              </p>
            </div>
            <div style={{ display: 'flex', gap: '6px', marginLeft: '10px', alignItems: 'center', flexShrink: 0 }}>
              {registro.Tipo_Apertura === 'Continuacion' && (
                <span style={{ backgroundColor: '#0277bd', borderRadius: '6px', padding: '2px 8px', fontSize: '10px', fontWeight: 700 }}>🔗 Continuación</span>
              )}
              <span style={{
                backgroundColor: registro.Estado === 'cerrado' ? '#1b5e20' : '#f57f17',
                borderRadius: '6px', padding: '2px 8px', fontSize: '10px', fontWeight: 700,
              }}>
                {registro.Estado === 'cerrado' ? '✓ Cerrado' : registro.Estado === 'transferido' ? '↔ Transferido' : '⏳ Abierto'}
              </span>
              <button onClick={onCerrar} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', borderRadius: '6px', color: 'white', cursor: 'pointer', fontSize: '14px', padding: '3px 8px' }}>✕</button>
            </div>
          </div>
        </div>

        {/* ── Semáforo — acción principal, arriba ── */}
        <div style={{ padding: '12px 16px', borderBottom: '1px solid #f0f0f0', flexShrink: 0, backgroundColor: '#fafafa' }}>
          <p style={{ fontSize: '11px', fontWeight: 700, color: '#555', textTransform: 'uppercase', margin: '0 0 8px' }}>Semáforo *</p>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '8px' }}>
            {['Verde', 'Amarillo', 'Rojo'].map((e) => (
              <button key={e} type="button" onClick={() => setEstado(e)} style={{
                padding: '10px', borderRadius: '10px', border: '2px solid',
                borderColor: estado === e ? COLORES_SEMAFORO[e] : '#e0e0e0',
                backgroundColor: estado === e ? BG_SEMAFORO[e] : 'white',
                color: COLORES_SEMAFORO[e], fontWeight: 700, fontSize: '14px',
                cursor: 'pointer', minHeight: '44px',
              }}>
                {estado === e ? '● ' : ''}{e}
              </button>
            ))}
          </div>
        </div>

        {/* ── Contenido scrollable ── */}
        <div style={{ overflowY: 'auto', flex: 1, padding: '14px 16px', display: 'flex', flexDirection: 'column', gap: '12px' }}>

          {/* Análisis calculado del turno completo */}
        <div>
          <p style={{ fontSize: '12px', fontWeight: 700, color: '#555', textTransform: 'uppercase', marginBottom: '8px' }}>
            Análisis del turno {esMultiProducto && <span style={{ color: '#F8A12F' }}>({registrosTurno.length} productos)</span>}
          </p>

          {/* Tabla de registros del turno si hay más de uno */}
          {esMultiProducto && (
            <div style={{ marginBottom: '10px', border: '1px solid #e0e0e0', borderRadius: '8px', overflow: 'hidden', fontSize: '12px' }}>
              <div style={{ backgroundColor: '#004895', color: 'white', padding: '7px 10px', fontWeight: 700, display: 'grid', gridTemplateColumns: '2fr 1fr 1fr 1fr 1fr', gap: '4px' }}>
                <span>Producto</span>
                <span style={{ textAlign:'right' }}>Decl.</span>
                <span style={{ textAlign:'right' }}>Teórico</span>
                <span style={{ textAlign:'right' }}>Def.</span>
                <span style={{ textAlign:'right' }}>Kg dev.</span>
              </div>
              {registrosTurno.map(r => {
                const esteReg = r.ID === registro.ID
                const abierto = r.Estado !== 'cerrado' && r.Estado !== 'transferido'
                const kpu = (kgPorUnidadMap || {})[r.Producto] || 0
                const mpCalc = r.MP_Consumida_Calc || 0
                const teorico = kpu > 0 && mpCalc > 0 ? Math.round(mpCalc / kpu) : null
                return (
                  <div key={r.ID} style={{ padding: '7px 10px', borderTop: '1px solid #f0f0f0', display: 'grid', gridTemplateColumns: '2fr 1fr 1fr 1fr 1fr', gap: '4px', backgroundColor: esteReg ? '#f0f4ff' : 'white' }}>
                    <span style={{ color: '#222', fontWeight: esteReg ? 700 : 400, fontSize: '12px' }}>
                      {r.Producto || '—'}
                      {abierto && <span style={{ marginLeft: '4px', color: '#f57f17', fontSize: '10px' }}>▶</span>}
                      {esteReg && <span style={{ marginLeft: '4px', color: '#004895', fontSize: '10px' }}>◄</span>}
                    </span>
                    <span style={{ textAlign:'right', color: '#2e7d32', fontWeight: 600 }}>
                      {abierto ? '—' : ((r.UnidadesConformes || 0) + (r.UnidadesDefectuosas || 0))}
                    </span>
                    <span style={{ textAlign:'right', color: '#555', fontStyle: teorico == null ? 'italic' : 'normal' }}>
                      {teorico != null ? teorico : '—'}
                    </span>
                    <span style={{ textAlign:'right', color: '#c62828' }}>{r.UnidadesDefectuosas ?? (abierto ? '—' : '0')}</span>
                    <span style={{ textAlign:'right', color: '#333' }}>{r.KgMPDevuelta != null ? r.KgMPDevuelta : (abierto ? '—' : '0')}</span>
                  </div>
                )
              })}
            </div>
          )}

          <div style={{ backgroundColor: (analisis.diferenciaPct||0) > 8 ? '#fff3f3' : '#f3fff3', borderRadius: '10px', padding: '12px', fontSize: '13px', border: `1px solid ${colorDif}40`, display: 'flex', flexDirection: 'column', gap: '6px' }}>
            {kgKardex != null && (
              <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                <span style={{ color: '#888', fontSize: '12px' }}>MP total OF (Kardex):</span>
                <span style={{ color: '#888', fontSize: '12px' }}>{kgKardex.toFixed(2)} kg</span>
              </div>
            )}
            <div style={{ display: 'flex', justifyContent: 'space-between' }}>
              <span style={{ color: '#555' }}>MP real Base (operario):</span>
              <strong style={{ color: '#111' }}>{(registro.MP_KgUsado ?? 0).toFixed(2)} kg</strong>
            </div>
            {(() => {
              // MP real: MP_KgUsado (real Base declarado por operario), editado por PCP si aplica
              const mpCons = parseFloat(mpLineaEdits.MP_KgUsado) || registro.MP_KgUsado || analisis.mpConsumida || 0
              // Producción declarada: usar edits si PCP los cambió, sino valores originales
              const confVal = edits.UnidadesConformes !== '' ? (parseInt(edits.UnidadesConformes) || 0) : (registro.UnidadesConformes || 0)
              const defVal  = edits.UnidadesDefectuosas !== '' ? (parseInt(edits.UnidadesDefectuosas) || 0) : (registro.UnidadesDefectuosas || 0)
              const prodDec = confVal + defVal
              // MP teórica = Producción declarada × KgPorUnidad
              const mpTeo   = kgPorUnidad > 0 ? prodDec * kgPorUnidad
                             : (registro.Produccion_Teorica != null ? Number(registro.Produccion_Teorica) : null)
              // Diferencia MP = |consumida - teórica| / consumida × 100
              const difPct  = mpCons > 0 && mpTeo != null
                             ? Math.abs(mpCons - mpTeo) / mpCons * 100
                             : (registro.Diferencia_Pct != null ? Number(registro.Diferencia_Pct) : null)
              const exceso  = mpCons > (mpTeo || 0)
              const colorD  = difPct == null ? '#555' : difPct > 8 ? '#c62828' : '#2e7d32'
              const prodTeo = kgPorUnidad > 0 && mpCons > 0 ? Math.round(mpCons / kgPorUnidad) : null
              return (
                <>
                  <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                    <span style={{ color: '#555' }}>MP teórica (Σ unid × kg/und):</span>
                    <strong style={{ color: '#111' }}>{mpTeo != null ? `${Number(mpTeo).toFixed(2)} kg` : '—'}</strong>
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: 700 }}>
                    <span style={{ color: colorD }}>Diferencia MP:</span>
                    <strong style={{ color: colorD }}>
                      {difPct != null ? `${exceso ? '↑' : '↓'} ${Number(difPct).toFixed(1)}% ${difPct > 8 ? '⚠' : '✓'}` : '—'}
                    </strong>
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', borderTop: '1px solid #ddd', paddingTop: '6px' }}>
                    <span style={{ color: '#555' }}>Producción declarada:</span>
                    <strong style={{ color: '#111' }}>{prodDec} und</strong>
                  </div>
                  {prodTeo != null && (
                    <>
                      <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                        <span style={{ color: '#555' }}>Producción teórica:</span>
                        <strong style={{ color: '#555' }}>{prodTeo} und</strong>
                      </div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: 700 }}>
                        <span style={{ color: Math.abs(prodDec - prodTeo) > prodTeo * 0.08 ? '#c62828' : '#2e7d32' }}>Diferencia producción:</span>
                        <strong style={{ color: Math.abs(prodDec - prodTeo) > prodTeo * 0.08 ? '#c62828' : '#2e7d32' }}>
                          {prodDec - prodTeo > 0 ? '+' : ''}{prodDec - prodTeo} und {Math.abs(prodDec - prodTeo) > prodTeo * 0.08 ? '⚠' : '✓'}
                        </strong>
                      </div>
                    </>
                  )}
                </>
              )
            })()}
          </div>
        </div>

        {/* Cruce Kardex vs Operario — MARCADOR POSICIÓN, se mueve abajo */}
        {false && <div>
          <p style={{ fontSize: '12px', fontWeight: 700, color: '#555', textTransform: 'uppercase', marginBottom: '8px' }}>Cruce Kardex MP</p>
          {kgKardex === null ? (
            <div style={{ backgroundColor: '#fff8e1', borderRadius: '10px', padding: '12px', border: '1px solid #ffcc02', fontSize: '13px', color: '#795548' }}>
              ⚠ Sin entrada en Kardex_MP para esta máquina / fecha / turno.
            </div>
          ) : (() => {
            const totKgOp = kgKardex ?? 0  // usar Kardex como referencia, no campo eliminado
            const diff = totKgOp - kgKardex
            const pct = kgKardex > 0 ? Math.abs(diff) / kgKardex * 100 : 0
            const ok = Math.abs(diff) < 0.5
            return (
              <div style={{ backgroundColor: ok ? '#e8f5e9' : '#ffebee', borderRadius: '10px', border: `1px solid ${ok ? '#4caf50' : '#f44336'}`, fontSize: '13px', overflow: 'hidden' }}>
                {/* Detalle de items Kardex */}
                {itemsKardexTurno.length > 0 && (
                  <div style={{ padding: '10px 12px', borderBottom: '1px solid rgba(0,0,0,0.08)' }}>
                    <p style={{ fontSize: '11px', fontWeight: 700, color: '#555', marginBottom: '6px', textTransform: 'uppercase' }}>Detalle Kardex</p>
                    {itemsKardexTurno.map(k => (
                      <div key={k.ID} style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', marginBottom: '3px' }}>
                        <span style={{ color: '#333' }}>{k.Insumo || '—'}</span>
                        <span style={{ color: '#1b5e20', fontWeight: 600 }}>
                          {k.KgEntregados} kg{k.KgDevueltos > 0 ? ` − ${k.KgDevueltos} dev = ${(k.KgEntregados - k.KgDevueltos).toFixed(2)} neto` : ''}
                        </span>
                      </div>
                    ))}
                  </div>
                )}
                <div style={{ padding: '10px 12px', display: 'flex', flexDirection: 'column', gap: '5px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                    <span style={{ color: '#555' }}>Total Kardex (neto):</span>
                    <strong style={{ color: '#111' }}>{kgKardex.toFixed(2)} kg</strong>
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                    <span style={{ color: '#555' }}>Total declarado por operario:</span>
                    <strong style={{ color: '#111' }}>{totKgOp.toFixed(2)} kg</strong>
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: 700, borderTop: '1px solid rgba(0,0,0,0.1)', paddingTop: '5px' }}>
                    <span style={{ color: ok ? '#2e7d32' : '#c62828' }}>Diferencia:</span>
                    <strong style={{ color: ok ? '#2e7d32' : '#c62828' }}>
                      {diff > 0 ? '+' : ''}{diff.toFixed(2)} kg ({pct.toFixed(1)}%) {ok ? '✓' : '⚠ REVISAR'}
                    </strong>
                  </div>
                </div>
              </div>
            )
          })()}
        </div>}

        {/* Datos del operario — editables directamente */}
        <div>
          <p style={{ fontSize: '12px', fontWeight: 700, color: '#555', textTransform: 'uppercase', margin: '0 0 8px' }}>
            Datos del operario <span style={{ fontSize: '10px', color: '#aaa', fontWeight: 400, textTransform: 'none' }}>— edita si hay corrección</span>
          </p>
          <div style={{ backgroundColor: 'white', borderRadius: '10px', border: '1.5px solid #e0e0e0', padding: '12px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
              {[
                ['UnidadesConformes',   'Und. conformes'],
                ['UnidadesDefectuosas', 'Und. defectuosas'],
                ['GruposConformes',     'Grupos conformes'],
                ['UnidadesSueltas',     'Und. sueltas'],
              ].map(([field, label]) => (
                <div key={field}>
                  <label style={{ fontSize: '11px', fontWeight: 600, color: '#666', display: 'block', marginBottom: '3px' }}>{label}</label>
                  <input
                    type="number" step="1" min="0"
                    value={edits[field]}
                    onChange={e => setEdits(prev => ({ ...prev, [field]: e.target.value }))}
                    style={{ width: '100%', padding: '8px 10px', borderRadius: '8px', border: '2px solid #ddd', fontSize: '16px', fontWeight: 700, textAlign: 'right', color: '#1a1a1a', backgroundColor: 'white', boxSizing: 'border-box' }}
                  />
                </div>
              ))}
            </div>
            <div style={{ fontSize: '11px', color: '#aaa', backgroundColor: '#f9f9f9', borderRadius: '6px', padding: '6px 10px' }}>
              💡 MP consumida se corrige en el panel <strong>Kardex MP</strong>
            </div>
            {/* Motivo solo si cambió algo */}
            {(() => {
              const cambio = Object.entries(edits).some(([k, v]) => v !== '' && Number(v) !== (registro[k] ?? 0))
              return cambio ? (
                <div>
                  <label style={{ fontSize: '11px', fontWeight: 600, color: '#c62828', display: 'block', marginBottom: '3px' }}>Motivo de la corrección *</label>
                  <input value={motivo} onChange={e => setMotivo(e.target.value)}
                    placeholder="¿Por qué se corrigen estos datos?"
                    style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '2px solid #f57f17', fontSize: '13px', color: '#1a1a1a', backgroundColor: 'white', boxSizing: 'border-box' }}
                  />
                </div>
              ) : null
            })()}
            {/* Gap 3: Autorización del Jefe para correcciones retroactivas */}
              {esRetroactivo && (
                <div style={{ backgroundColor: '#ffebee', borderRadius: '8px', padding: '10px 12px', border: '1px solid #ef9a9a' }}>
                  <p style={{ fontSize: '12px', color: '#c62828', fontWeight: 700, margin: '0 0 8px' }}>
                    ⚠ Corrección retroactiva — registro del {fechaReg}
                  </p>
                  <p style={{ fontSize: '11px', color: '#555', margin: '0 0 8px' }}>
                    Según el manual, debes obtener autorización del Jefe de Operaciones antes de corregir datos de días anteriores.
                  </p>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' }}>
                    <input type="checkbox" checked={autorizadoPorJefe} onChange={e => setAutorizadoPorJefe(e.target.checked)}
                      style={{ width: '16px', height: '16px', accentColor: '#c62828' }} />
                    <span style={{ fontSize: '12px', fontWeight: 600, color: '#c62828' }}>
                      Confirmo que el Jefe de Operaciones autorizó esta corrección
                    </span>
                  </label>
                </div>
              )}
          </div>
        </div>

        {/* MP_KgUsado — solo campo real Base MP, editable por PCP */}
        <div>
          <p style={{ fontSize: '12px', fontWeight: 700, color: '#555', textTransform: 'uppercase', margin: '0 0 8px' }}>
            MP real Base <span style={{ fontSize: '10px', color: '#aaa', fontWeight: 400, textTransform: 'none' }}>— declarado por operario, verificar en Kardex</span>
          </p>
          <div style={{ backgroundColor: 'white', borderRadius: '10px', border: '1.5px solid #c5cae9', padding: '10px 12px' }}>
            <label style={{ fontSize: '11px', fontWeight: 700, color: '#004895', display: 'block', marginBottom: '4px' }}>
              Kg usado en producción
              {(registro.Insumo_Base || itemsBase.map(k => k.Insumo).join(', ')) && (
                <span style={{ marginLeft: '6px', color: '#555', fontWeight: 600 }}>
                  ({registro.Insumo_Base || itemsBase.map(k => k.Insumo).join(', ')})
                </span>
              )}
            </label>
            <input type="number" step="0.01" min="0"
              value={mpLineaEdits.MP_KgUsado ?? ''}
              onChange={ev => setMpLineaEdits({ MP_KgUsado: ev.target.value })}
              placeholder="0"
              style={{ width: '100%', padding: '8px 10px', borderRadius: '8px', border: '2px solid #004895a0', fontSize: '20px', fontWeight: 800, textAlign: 'right', color: '#1a1a1a', backgroundColor: 'white', boxSizing: 'border-box' }}
            />
            <p style={{ fontSize: '10px', color: '#999', margin: '4px 0 0' }}>
              Merma y devoluciones → editar en Kardex MP
            </p>
          </div>
        </div>

        {/* Obs PCP */}
        <div>
          <label style={{ ...LABEL, marginBottom: '4px' }}>Observación PCP</label>
          <textarea value={obs} onChange={e => setObs(e.target.value)} rows={2}
            placeholder="Escribe tu observación (opcional)..."
            style={{ ...INPUT_STYLE, resize: 'none', fontSize: '13px', padding: '8px 10px' }}
          />
        </div>

        </div> {/* fin contenido scrollable */}

        {/* ── Footer sticky ── */}
        <div style={{ padding: '12px 16px', borderTop: '1px solid #e0e0e0', backgroundColor: 'white', flexShrink: 0 }}>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', gap: '8px' }}>
            <button onClick={onCerrar} style={{
              padding: '13px', borderRadius: '10px', border: '2px solid #ddd',
              backgroundColor: 'white', color: '#555', fontSize: '14px', cursor: 'pointer',
            }}>Cancelar</button>
            <button onClick={() => onGuardar({ estado, obs, motivo, editando, edits, autorizadoPorJefe, kardexEdits, mpLineaEdits })} disabled={guardando} style={{
              padding: '13px', borderRadius: '10px', border: 'none',
              backgroundColor: guardando ? '#ccc' : COLORES_SEMAFORO[estado] || '#004895',
              color: 'white', fontSize: '15px', fontWeight: 700,
              cursor: guardando ? 'not-allowed' : 'pointer',
            }}>
              {guardando ? '⏳ Guardando...' : `✓ Guardar como ${estado}`}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}

// ─── Pantalla principal ───────────────────────────────────────────────────────
export default function ValidacionPCP({ onIrKardex, onIrTablets, onLogout, onCambiarRol }) {
  const { getToken, usuario } = useMsal()
  const [registros, setRegistros] = useState([])
  const [productos, setProductos] = useState({})       // { codigo: kgPorUnidad }
  const [nombresProductos, setNombresProductos] = useState({}) // { codigo: nombre }
  const [tiposInsumo, setTiposInsumo] = useState({})    // { nombre: 'Base' | 'Colorante' }
  const [kardexMap, setKardexMap]     = useState({})
  const [kardexItems, setKardexItems] = useState([]) // raw para detalle en modal
  const [cargando, setCargando] = useState(true)
  const [filtroFecha, setFiltroFecha] = useState(format(new Date(), 'yyyy-MM-dd'))
  const [filtroEstado, setFiltroEstado] = useState('todos')
  const [modalReg, setModalReg] = useState(null)
  const [guardando, setGuardando] = useState(false)

  const cargar = useCallback(async () => {
    setCargando(true)
    try {
      const token = await getToken()
      if (!token) return
      const [regs, prods, kardexItems] = await Promise.allSettled([
        getListItems(token, 'Registro_Produccion', { orderby: 'HoraInicio desc', top: 200 }),
        getListItems(token, 'Maestro_Productos'),
        getListItems(token, 'Kardex_MP', { top: 500 }),
      ])

      const prodList = prods.status === 'fulfilled' ? prods.value : []
      const mapProd = {}
      const mapNombres = {}
      const tiposMap = {}
      prodList.forEach(p => {
        const nombre = p.Nombre || p.Title || ''
        // Indexar por código Y por nombre para tolerar que Registro_Produccion guarde cualquiera
        const tipo = (p.TipoProducto || '').toUpperCase()
        const esTipo = tipo === 'MC' ? 'Colorante' : 'Base'
        if (p.Codigo) {
          mapProd[p.Codigo]  = p.KgPorUnidad
          mapNombres[p.Codigo] = nombre
        }
        if (nombre) {
          mapProd[nombre]    = mapProd[nombre] ?? p.KgPorUnidad
          mapNombres[nombre] = nombre
        }
        // Indexar tipo por nombre para filtrar Kardex en el modal
        tiposMap[nombre.toLowerCase()] = esTipo
        if (p.Codigo) tiposMap[p.Codigo.toLowerCase()] = esTipo
      })
      setProductos(mapProd)
      setNombresProductos(mapNombres)
      setTiposInsumo(tiposMap)

      if (regs.status === 'fulfilled') setRegistros(regs.value)

      // Mapa kardex: "Maquina|Fecha|Turno" → KgEntregados
      if (kardexItems.status === 'fulfilled') {
        const raw = kardexItems.value
        setKardexItems(raw)
        const map = {}
        raw.forEach(k => {
          const fecha = k.Fecha ? k.Fecha.split('T')[0] : ''
          const maquina = k.Title || k.Maquina || ''
          const of = k.Numero_OF || ''
          const key = `${maquina}|${of}|${fecha}|${k.Turno}`
          const neto = (k.KgEntregados || 0) - (k.KgDevueltos || 0)
          map[key] = parseFloat(((map[key] || 0) + neto).toFixed(3))
        })
        setKardexMap(map)
      }
    } catch (err) {
      console.error('Error cargando registros PCP:', err)
    } finally {
      setCargando(false)
    }
  }, [])

  useEffect(() => { cargar() }, [cargar])

  const registrosFiltrados = registros.filter(r => {
    const fechaReg = r.Fecha ? r.Fecha.split('T')[0] : (r.HoraInicio ? r.HoraInicio.split('T')[0] : '')
    const matchFecha = !filtroFecha || fechaReg === filtroFecha
    const matchEstado = filtroEstado === 'todos' || (r.Estado_Validacion || 'Sin semáforo') === filtroEstado
    return matchFecha && matchEstado
  })

  const handleGuardar = async ({ estado, obs, motivo, editando, edits, autorizadoPorJefe, kardexEdits, mpLineaEdits }) => {
    // Detectar si realmente hubo cambios en los datos
    const huboCambios = editando && Object.entries(edits).some(([k, v]) =>
      v !== '' && Number(v) !== (modalReg?.[k] ?? 0)
    )

    if (huboCambios && !motivo.trim()) {
      alert('El motivo de corrección es obligatorio cuando se editan datos.')
      return
    }

    // Gap 3: Corrección retroactiva requiere autorización del Jefe
    if (huboCambios) {
      const fechaReg = modalReg.Fecha ? modalReg.Fecha.split('T')[0] : ''
      const hoy = new Date().toISOString().split('T')[0]
      if (fechaReg && fechaReg < hoy && !autorizadoPorJefe) {
        alert('Este registro es de un día anterior. Obtén autorización del Jefe de Operaciones antes de corregir.')
        return
      }
    }

    setGuardando(true)
    try {
      const token = await getToken()

      // Campos de validación (siempre)
      const payload = {
        Estado_Validacion: estado,
        Fecha_Validacion: new Date().toISOString(),
      }
      // Obs_PCP solo si tiene valor (evita error en columnas Multiple Lines con Rich Text)
      if (obs && obs.trim()) payload.Obs_PCP = obs

      // Si PCP editó datos del operario: guardar + recalcular campos derivados
      if (huboCambios && motivo.trim()) {
        const originales = {
          UnidadesConformes:   modalReg.UnidadesConformes,
          UnidadesDefectuosas: modalReg.UnidadesDefectuosas,
          GruposConformes:     modalReg.GruposConformes,
          UnidadesSueltas:     modalReg.UnidadesSueltas,
        }
        payload.Valor_Corregido = `Valores originales: ${JSON.stringify(originales)}`
        payload.Motivo_Correccion = motivo
        // Gap 1: Corregido_Por como Person/Group — requiere ID de usuario SP
        try {
          const userId = await resolveUserId(token, usuario?.email || '')
          if (userId) payload.Corregido_PorId = userId
        } catch { /* sin permiso para resolver usuario — omitir campo */ }

        if (edits.UnidadesConformes   !== '')  payload.UnidadesConformes   = parseInt(edits.UnidadesConformes,   10)
        if (edits.UnidadesDefectuosas !== '')  payload.UnidadesDefectuosas = parseInt(edits.UnidadesDefectuosas, 10)
        if (edits.GruposConformes     !== '')  payload.GruposConformes     = parseInt(edits.GruposConformes,     10)
        if (edits.UnidadesSueltas     !== '')  payload.UnidadesSueltas     = parseInt(edits.UnidadesSueltas,     10)
      }

      // ── MP_KgUsado: guardar en payload y actualizar Kardex_MP ────────────────
      if (mpLineaEdits?.MP_KgUsado !== undefined) {
        const kgUsadoNuevo = parseFloat(mpLineaEdits.MP_KgUsado) || 0
        payload.MP_KgUsado = kgUsadoNuevo

        if (modalReg.Numero_OF) {
          // Insumos Base que usó este registro (guardados en Insumo_Base al cierre)
          const insumosBase = (modalReg.Insumo_Base || '')
            .split(',').map(s => s.trim().toLowerCase()).filter(Boolean)

          // Para cada insumo Base de este registro, calcular su acumulado y actualizar Kardex
          const insumosAActualizar = insumosBase.length > 0
            ? insumosBase
            : kardexItems
                .filter(k => (k.Numero_OF || '') === modalReg.Numero_OF &&
                             (tiposInsumo[(k.Insumo || '').toLowerCase()] || 'Base') !== 'Colorante')
                .map(k => (k.Insumo || '').toLowerCase())

          for (const insumoNombre of [...new Set(insumosAActualizar)]) {
            // Acumulado para este insumo específico:
            // suma MP_KgUsado de todos los registros que usaron este insumo
            const acumInsumo = registros
              .filter(r => r.Numero_OF === modalReg.Numero_OF &&
                           (r.Insumo_Base || '').toLowerCase().includes(insumoNombre))
              .reduce((s, r) => s + (r.ID === modalReg.ID ? kgUsadoNuevo : (r.MP_KgUsado || 0)), 0)

            const entradaKardex = kardexItems.find(k =>
              (k.Numero_OF || '') === modalReg.Numero_OF &&
              (k.Insumo || '').toLowerCase() === insumoNombre
            )
            if (entradaKardex) {
              const nuevoKgUsado = parseFloat(acumInsumo.toFixed(3))
              await updateListItem(token, 'Kardex_MP', entradaKardex.ID, { KgUsado: nuevoKgUsado })
              setKardexItems(prev => prev.map(ki =>
                ki.ID === entradaKardex.ID ? { ...ki, KgUsado: nuevoKgUsado } : ki
              ))
            }
          }
        }
      }

      // ── Recálculo consolidado de campos derivados ────────────────────────────
      {
        const kgPorUnidad  = productos[modalReg.Producto] ?? 0
        const nuevoConf    = payload.UnidadesConformes   ?? modalReg.UnidadesConformes   ?? 0
        const nuevoDef     = payload.UnidadesDefectuosas ?? modalReg.UnidadesDefectuosas ?? 0
        const prodTotal    = nuevoConf + nuevoDef
        const mpKgUsado    = payload.MP_KgUsado ?? modalReg.MP_KgUsado ?? 0

        if (kgPorUnidad > 0) {
          // MP teórica según unidades producidas
          payload.MP_Consumida_Calc  = parseFloat((prodTotal * kgPorUnidad).toFixed(3))
          payload.Produccion_Teorica = prodTotal
        }

        if (kgPorUnidad > 0 && mpKgUsado > 0) {
          const mpTeorica     = prodTotal * kgPorUnidad
          const diferenciaPct = parseFloat((Math.abs(mpKgUsado - mpTeorica) / mpKgUsado * 100).toFixed(2))
          payload.Diferencia_Pct = diferenciaPct
        }

        // KgMPRestante = total entregado Kardex OF - KgUsado acumulado actualizado
        if (modalReg.Numero_OF) {
          const kardexOF = kardexItems.filter(k => (k.Numero_OF || '') === modalReg.Numero_OF)
          const totalEntregado = kardexOF.reduce((s, k) => s + (k.KgEntregados || 0), 0)
          const totalUsado     = kardexOF.reduce((s, k) => s + (k.KgUsado || 0), 0)
          payload.KgMPRestante = parseFloat(Math.max(0, totalEntregado - totalUsado).toFixed(3))
        }
      }

      await updateListItem(token, 'Registro_Produccion', modalReg.ID, payload)
      // Actualizar registro en memoria
      setRegistros(prev => prev.map(r => r.ID === modalReg.ID ? { ...r, ...payload } : r))

      // Guardar Merma/Devueltos de Kardex si PCP los editó directamente (flujo anterior)
      if (kardexEdits && Object.keys(kardexEdits).length > 0) {
        await Promise.allSettled(
          Object.entries(kardexEdits).map(([id, e]) =>
            updateListItem(token, 'Kardex_MP', parseInt(id), {
              KgMermaRec:   parseFloat(e.KgMermaRec)   || 0,
              KgMermaNoRec: parseFloat(e.KgMermaNoRec) || 0,
              KgDevueltos:  parseFloat(e.KgDevueltos)  || 0,
            })
          )
        )
      }

      setModalReg(null)
    } catch (err) {
      alert('Error al guardar: ' + err.message)
    } finally {
      setGuardando(false)
    }
  }

  const [gruposAbiertos, setGruposAbiertos] = useState({})
  const toggleGrupo = (key) => setGruposAbiertos(prev => ({ ...prev, [key]: !prev[key] }))

  // Estado de cada OF = el peor estado entre sus registros (Rojo > Amarillo > Verde)
  const estadosPorOf = (() => {
    const rank = { 'Rojo': 3, 'Amarillo': 2, 'Verde': 1, 'Sin semáforo': 0 }
    const map = {}
    registros.forEach(r => {
      const key = `${r.Title || ''}|${r.Numero_OF || 'sin-of'}`
      const est = r.Estado_Validacion || 'Sin semáforo'
      if (!map[key] || (rank[est] || 0) > (rank[map[key]] || 0)) map[key] = est
    })
    return Object.values(map)
  })()
  const countPorEstado = (e) => estadosPorOf.filter(s => s === e).length

  // Jerarquía: OF → Turno → Productos
  const ofGroups = (() => {
    const filtrados = registros.filter(r => {
      const fechaReg = (r.Fecha || r.HoraInicio || '').split('T')[0]
      const matchFecha = !filtroFecha || fechaReg === filtroFecha
      const matchEstado = filtroEstado === 'todos' || (r.Estado_Validacion || 'Sin semáforo') === filtroEstado
      return matchFecha && matchEstado
    })

    const ofMap = {}
    filtrados.forEach(r => {
      const ofKey = `${r.Title || ''}|${r.Numero_OF || 'sin-of'}`
      if (!ofMap[ofKey]) {
        ofMap[ofKey] = {
          ofKey, maquina: r.Title || '—', numeroOF: r.Numero_OF || null,
          turnos: {},
        }
      }
      const tk = r.Turno || 'M'
      if (!ofMap[ofKey].turnos[tk]) {
        const fecha = (r.Fecha || r.HoraInicio || '').split('T')[0]
        ofMap[ofKey].turnos[tk] = { turno: tk, fecha, items: [] }
      }
      ofMap[ofKey].turnos[tk].items.push(r)
    })
    return Object.values(ofMap)
  })()

  return (
    <div style={{ backgroundColor: '#f5f5f5', minHeight: '100vh' }}>
      {/* Header PCP */}
      <header style={{ backgroundColor: '#37BEEC', color: 'white', padding: '12px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', position: 'sticky', top: 0, zIndex: 100 }}>
        <div>
          <p style={{ fontSize: '11px', opacity: 0.7, textTransform: 'uppercase' }}>Panel PCP</p>
          <h1 style={{ fontSize: '17px', fontWeight: 700 }}>Validación de registros</h1>
        </div>
        <div style={{ display: 'flex', gap: '8px' }}>
          <button onClick={onIrKardex} style={{ backgroundColor: 'rgba(255,255,255,0.25)', color: 'white', border: '1.5px solid rgba(255,255,255,0.6)', borderRadius: '8px', padding: '8px 14px', fontSize: '13px', fontWeight: 600, cursor: 'pointer' }}>
            📦 Kardex MP
          </button>
          {onIrTablets && (
            <button onClick={onIrTablets} style={{ backgroundColor: 'rgba(255,255,255,0.15)', color: 'white', border: '1.5px solid rgba(255,255,255,0.4)', borderRadius: '8px', padding: '6px 12px', fontSize: '12px', cursor: 'pointer' }}>
              🖥 Tablets
            </button>
          )}
          <button onClick={onCambiarRol} style={{ backgroundColor: 'transparent', border: '1.5px solid rgba(255,255,255,0.5)', color: 'white', borderRadius: '8px', padding: '6px 12px', fontSize: '12px', cursor: 'pointer' }}>
            Cambiar rol
          </button>
          <button onClick={onLogout} style={{ backgroundColor: 'transparent', border: '1.5px solid rgba(255,255,255,0.5)', color: 'white', borderRadius: '8px', padding: '6px 12px', fontSize: '12px', cursor: 'pointer' }}>
            Salir
          </button>
        </div>
      </header>

      <div style={{ padding: '16px', maxWidth: '900px', margin: '0 auto' }}>

        {/* KPIs semáforo */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '10px', marginBottom: '16px' }}>
          {[['Rojo', '🔴'], ['Amarillo', '🟡'], ['Verde', '🟢']].map(([e, icon]) => (
            <div key={e} onClick={() => setFiltroEstado(filtroEstado === e ? 'todos' : e)}
              style={{ backgroundColor: filtroEstado === e ? BG_SEMAFORO[e] : 'white', borderRadius: '12px', padding: '14px', textAlign: 'center', cursor: 'pointer', border: `2px solid ${filtroEstado === e ? COLORES_SEMAFORO[e] : '#ddd'}`, boxShadow: '0 1px 4px rgba(0,0,0,0.08)' }}>
              <div style={{ fontSize: '24px', fontWeight: 800, color: COLORES_SEMAFORO[e] }}>{countPorEstado(e)}</div>
              <div style={{ fontSize: '12px', color: '#666' }}>{icon} {e}s</div>
            </div>
          ))}
        </div>

        {/* Filtros */}
        <div style={{ display: 'flex', gap: '8px', marginBottom: '16px', alignItems: 'center' }}>
          <input type="date" value={filtroFecha} onChange={e => setFiltroFecha(e.target.value)}
            style={{ flex: '1 1 0', minWidth: 0, padding: '9px 8px', borderRadius: '8px', border: '2px solid #ddd', fontSize: '13px', backgroundColor: 'white', color: '#1a1a1a' }} />
          <select value={filtroEstado} onChange={e => setFiltroEstado(e.target.value)}
            style={{ padding: '9px 6px', borderRadius: '8px', border: '2px solid #ddd', fontSize: '13px', backgroundColor: 'white', color: '#1a1a1a', width: '96px', flexShrink: 0 }}>
            <option value="todos">Todos</option>
            <option value="Rojo">Rojo</option>
            <option value="Amarillo">Amarillo</option>
            <option value="Verde">Verde</option>
            <option value="Sin semáforo">Sin semáforo</option>
          </select>
          <button onClick={() => { setFiltroFecha(''); setFiltroEstado('todos') }}
            style={{ padding: '9px 10px', borderRadius: '8px', border: '1.5px solid #ddd', backgroundColor: 'white', color: '#555', fontSize: '12px', cursor: 'pointer', whiteSpace: 'nowrap', flexShrink: 0 }}>
            Ver todo
          </button>
        </div>

        {/* Jerarquía: OF → Turno → Productos */}
        {cargando ? (
          <div style={{ textAlign: 'center', padding: '40px', color: '#666' }}>Cargando registros...</div>
        ) : ofGroups.length === 0 ? (
          <div style={{ textAlign: 'center', padding: '40px', color: '#888', backgroundColor: 'white', borderRadius: '12px' }}>
            No hay registros para los filtros seleccionados.
          </div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
            {ofGroups.map(ofGroup => {
              const ofKey = ofGroup.ofKey
              const abierto = gruposAbiertos[ofKey] === true
              const todosLosItems = Object.values(ofGroup.turnos).flatMap(t => t.items)
              const kgKardexOf = Object.values(ofGroup.turnos).reduce((s, t) => {
                const k = kardexMap[`${ofGroup.maquina}|${ofGroup.numeroOF || ''}|${t.fecha}|${t.turno}`] ?? 0
                return s + k
              }, 0)
              // Total MP realmente usada en la OF = suma de MP_KgUsado (declarado por operario)
              const kgUsadoOf = todosLosItems
                .filter(r => r.Estado === 'cerrado')
                .reduce((s, r) => s + (r.MP_KgUsado || 0), 0)
              const estadoOf = todosLosItems.some(r => r.Estado_Validacion === 'Rojo') ? 'Rojo'
                : todosLosItems.some(r => r.Estado_Validacion === 'Amarillo') ? 'Amarillo'
                : todosLosItems.every(r => r.Estado_Validacion === 'Verde') ? 'Verde'
                : null
              const cerradosOf = todosLosItems.filter(r => r.Estado === 'cerrado')
              // "Validado PCP" solo cuando PCP revisó todos los registros Y ninguno sigue en Rojo
              const todosValidados = cerradosOf.length > 0
                && cerradosOf.every(r => r.Fecha_Validacion)
                && !cerradosOf.some(r => r.Estado_Validacion === 'Rojo')

              return (
                <div key={ofKey} style={{ borderRadius: '12px', border: `2px solid ${COLORES_SEMAFORO[estadoOf] || '#e0e4ea'}`, overflow: 'hidden', boxShadow: '0 1px 6px rgba(0,0,0,0.07)' }}>

                  {/* Cabecera OF */}
                  <div onClick={() => toggleGrupo(ofKey)} style={{
                    backgroundColor: estadoOf ? BG_SEMAFORO[estadoOf] : '#f8f9fa',
                    padding: '10px 14px', cursor: 'pointer',
                    display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '8px',
                  }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      {/* Fila 1: semáforo + máquina + productos */}
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
                        {estadoOf && <Badge estado={estadoOf} />}
                        <strong style={{ fontSize: '14px', color: '#1a1a1a' }}>{ofGroup.maquina}</strong>
                        <span style={{ fontSize: '11px', color: '#666', backgroundColor: 'rgba(0,0,0,0.06)', borderRadius: '6px', padding: '1px 6px' }}>
                          {todosLosItems.length} prod.
                        </span>
                        {todosValidados && (
                          <span style={{ backgroundColor: '#e8f5e9', color: '#2e7d32', border: '1px solid #4caf50', borderRadius: '8px', padding: '1px 7px', fontSize: '10px', fontWeight: 700 }}>
                            ✓ PCP
                          </span>
                        )}
                      </div>
                      {/* Fila 2: OF + kg */}
                      {ofGroup.numeroOF && (
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginTop: '3px' }}>
                          <span style={{ fontSize: '11px', fontWeight: 700, color: '#444', fontFamily: 'monospace' }}>{ofGroup.numeroOF}</span>
                          {(kgUsadoOf > 0 || kgKardexOf > 0) && (
                            <span style={{ fontSize: '11px', color: '#777' }}>
                              {kgUsadoOf > 0 && <span style={{ color: '#1b5e20', fontWeight: 700 }}>{kgUsadoOf.toFixed(1)} kg usado</span>}
                              {kgKardexOf > 0 && <span>{kgUsadoOf > 0 ? ' / ' : ''}{kgKardexOf.toFixed(1)} entregado</span>}
                            </span>
                          )}
                        </div>
                      )}
                    </div>
                    <span style={{ fontSize: '14px', color: '#888', flexShrink: 0 }}>{abierto ? '▲' : '▼'}</span>
                  </div>

                  {/* Turnos dentro de la OF */}
                  {abierto && Object.values(ofGroup.turnos).map((turnoGrupo, tIdx) => {
                    const kgKardexTurno = kardexMap[`${ofGroup.maquina}|${ofGroup.numeroOF || ''}|${turnoGrupo.fecha}|${turnoGrupo.turno}`] ?? null
                    const analisis = calcularSemaforoTurno(turnoGrupo.items, productos, kgKardexTurno)

                    return (
                      <div key={turnoGrupo.turno} style={{ borderTop: tIdx > 0 ? '2px solid #e0e4ea' : '1px solid #e0e4ea' }}>
                        {/* Sub-cabecera Turno */}
                        <div style={{ backgroundColor: '#f8f9fa', padding: '8px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                            <span style={{ fontWeight: 700, fontSize: '13px', color: '#004895' }}>
                              Turno {turnoGrupo.turno === 'M' ? 'Mañana' : turnoGrupo.turno === 'T' ? 'Tarde' : 'Noche'}
                            </span>
                            <span style={{ fontSize: '11px', color: '#888' }}>
                              {turnoGrupo.fecha ? format(new Date(turnoGrupo.fecha + 'T12:00:00'), 'dd/MM/yyyy') : '—'}
                            </span>
                          </div>
                          <span style={{ fontSize: '11px', color: '#888' }}>
                            {turnoGrupo.items.length} registro{turnoGrupo.items.length > 1 ? 's' : ''}
                          </span>
                        </div>

                        {/* Productos del turno */}
                        {turnoGrupo.items.map((reg, idx) => {
                          const estadoReg = reg.Estado_Validacion || null
                          return (
                            <div key={reg.ID} style={{
                              padding: '10px 16px',
                              borderTop: '1px solid #f0f0f0',
                              display: 'grid', gridTemplateColumns: '12px 1fr auto',
                              gap: '10px', alignItems: 'center',
                            }}>
                              <div style={{ width: '10px', height: '10px', borderRadius: '50%', backgroundColor: COLORES_SEMAFORO[estadoReg] || '#ccc', flexShrink: 0 }} title={estadoReg || 'Sin validar'} />
                              <div>
                                <p style={{ fontSize: '13px', fontWeight: 700, color: '#1a1a1a', margin: 0 }}>
                                  {nombresProductos[reg.Producto] || reg.Producto || '—'}
                                  {reg.Estado === 'abierto' && <span style={{ backgroundColor: '#e3f2fd', color: '#1565c0', borderRadius: '6px', padding: '1px 7px', fontSize: '10px', marginLeft: '6px' }}>ABIERTO</span>}
                                </p>
                                {reg.Color && (
                                  <p style={{ fontSize: '12px', color: '#555', margin: 0, marginTop: '2px' }}>
                                    🎨 {reg.Color}
                                  </p>
                                )}
                                <p style={{ fontSize: '11px', color: '#888', margin: 0, marginTop: '2px' }}>
                                  👤 {reg.Operario || '—'} &nbsp;·&nbsp; ✓ {reg.UnidadesConformes ?? '—'} &nbsp;✗ {reg.UnidadesDefectuosas ?? '—'}
                                </p>
                              </div>
                              <button onClick={() => setModalReg(reg)} style={{
                                backgroundColor: '#004895', color: 'white', border: 'none',
                                borderRadius: '8px', padding: '7px 12px', fontSize: '12px',
                                fontWeight: 600, cursor: 'pointer', whiteSpace: 'nowrap',
                              }}>Validar</button>
                            </div>
                          )
                        })}
                      </div>
                    )
                  })}
                </div>
              )
            })}
          </div>
        )}
      </div>

      {modalReg && (() => {
        const key = turnoKey(modalReg)
        const registrosTurno = registros.filter(r => turnoKey(r) === key)
        const fecha = (modalReg.Fecha || modalReg.HoraInicio || '').split('T')[0]
        const kgKardex = kardexMap[`${modalReg.Title}|${modalReg.Numero_OF || ''}|${fecha}|${modalReg.Turno}`] ?? null
        const itemsKardexTurno = kardexItems.filter(k => {
          const kFecha = k.Fecha ? k.Fecha.split('T')[0] : ''
          return (k.Title || '') === modalReg.Title
            && kFecha === fecha
            && k.Turno === modalReg.Turno
            && (k.Numero_OF || '') === (modalReg.Numero_OF || '')
        })
        const registroConNombre = {
          ...modalReg,
          NombreProducto: nombresProductos[modalReg.Producto] || modalReg.Producto,
        }
        return (
          <ModalCorreccion
            registro={registroConNombre}
            kgPorUnidad={productos[modalReg.Producto] ?? 0}
            kgPorUnidadMap={productos}
            tiposInsumo={tiposInsumo}
            kgKardex={kgKardex}
            itemsKardexTurno={itemsKardexTurno}
            registrosTurno={registrosTurno}
            onGuardar={handleGuardar}
            onCerrar={() => setModalReg(null)}
            guardando={guardando}
          />
        )
      })()}
    </div>
  )
}
