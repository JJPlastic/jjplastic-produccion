import { useState, useEffect, useCallback } from 'react'
import { format, parseISO } from 'date-fns'
import { SearchSelect } from '../../components/SearchSelect'
import { useMsal } from '../../hooks/useMsal'
import { getListItems, createListItem, updateListItem, getOFsActivas } from '../../services/sharepoint'

const TURNOS = [{ id: 'M', label: 'Mañana' }, { id: 'T', label: 'Tarde' }, { id: 'N', label: 'Noche' }]

const sel = {
  width: '100%', padding: '11px', borderRadius: '8px',
  border: '2px solid #ddd', fontSize: '14px',
  backgroundColor: 'white', color: '#1a1a1a',
}
const inp = { ...sel }
const lbl = { fontSize: '13px', fontWeight: 600, color: '#333', display: 'block', marginBottom: '4px' }

const filaVacia = () => ({ id: crypto.randomUUID(), mp: '', kg: '' })

export default function KardexMP({ onVolver, onLogout }) {
  const { getToken } = useMsal()

  // ── Tab activa ────────────────────────────────────────────────────────────
  const [tab, setTab] = useState('entregas') // 'entregas' | 'registrar'

  // ── Campos base compartidos ────────────────────────────────────────────────
  const [fecha, setFecha] = useState(format(new Date(), 'yyyy-MM-dd'))
  const [turno, setTurno] = useState(() => {
    const h = new Date().getHours()
    return h >= 6 && h < 14 ? 'M' : h >= 14 && h < 22 ? 'T' : 'N'
  })
  const [maquina, setMaquina]       = useState('')
  const [obs, setObs]               = useState('')
  // Campos de orden de producción (integrados en Kardex)
  const [codigoLote, setCodigoLote] = useState('')
  const [producto, setProducto]     = useState('')
  const [cantPlan, setCantPlan]     = useState('')
  const [fechaFinEst, setFechaFinEst] = useState('')
  const [lotesActivos, setLotesActivos] = useState([])
  const [ofsActivasMaq, setOfsActivasMaq] = useState([]) // OFs activas cuando se selecciona máquina

  // ── Filas dinámicas de MP ─────────────────────────────────────────────────
  const [filas, setFilas] = useState([filaVacia()])

  const agregarFila  = () => setFilas(prev => [...prev, filaVacia()])
  const eliminarFila = (id) => setFilas(prev => prev.length > 1 ? prev.filter(f => f.id !== id) : prev)
  const updateFila   = (id, campo, valor) =>
    setFilas(prev => prev.map(f => f.id === id ? { ...f, [campo]: valor } : f))

  // ── Datos ─────────────────────────────────────────────────────────────────
  const [maquinas, setMaquinas]             = useState([])
  const [materiasPrimas, setMateriasPrimas] = useState([])
  const [entradas, setEntradas]             = useState([])
  const [cargando, setCargando]             = useState(true)
  const [enviando, setEnviando]             = useState(false)
  const [feedback, setFeedback]             = useState(null)
  // Historial: filtro y grupos
  const [filtroFechaHist, setFiltroFechaHist] = useState(format(new Date(), 'yyyy-MM-dd'))
  const [filtroTurnoHist, setFiltroTurnoHist] = useState('todos')
  const [gruposAbiertos, setGruposAbiertos]   = useState({})
  const toggleGrupo = (key) => setGruposAbiertos(p => ({ ...p, [key]: !p[key] }))
  // Modal agregar MP a grupo existente
  const [modalAdicional, setModalAdicional]   = useState(null) // { maquina, turno, fecha }
  const [filasAdicionales, setFilasAdicionales] = useState([filaVacia()])
  const [guardandoAd, setGuardandoAd]         = useState(false)
  // Devolución inline
  const [editDevId, setEditDevId]             = useState(null)
  const [editDevKg, setEditDevKg]             = useState('')
  const [guardandoDev, setGuardandoDev]       = useState(false)
  // Edición completa de una entrada (insumo, kg, mermas)
  const [editEntradaId, setEditEntradaId]     = useState(null)
  const [editEntrada, setEditEntrada]         = useState({})
  const [guardandoEntrada, setGuardandoEntrada] = useState(false)

  const cargar = useCallback(async () => {
    setCargando(true)
    try {
      const token = await getToken()
      if (!token) return
      const [maqRes, prodsRes, kardexRes] = await Promise.allSettled([
        getListItems(token, 'Maestro_Maquinas'),
        getListItems(token, 'Maestro_Productos', { top: 500 }),  // sin top=500 pierde los MP
        getListItems(token, 'Kardex_MP', { orderby: 'Created desc', top: 200 }),
      ])
      if (maqRes.status === 'fulfilled') {
        const activas = maqRes.value.filter(m =>
          m.Activo === true || m.Activo === 1 || m.Activo === undefined
        )
        setMaquinas(activas)
      }
      if (prodsRes.status === 'fulfilled') {
        const esActivo = p => p.Activo === true || p.Activo === 1 || p.Activo === undefined
        setMateriasPrimas(prodsRes.value.filter(p =>
          esActivo(p) && ['MP', 'MC'].includes((p.TipoProducto || '').toUpperCase())
        ))
      }
      if (kardexRes.status === 'fulfilled') setEntradas(kardexRes.value)

      // Lotes activos: registros de producción abiertos con su Codigo_Lote
      try {
        const regs = await getListItems(token, 'Registro_Produccion', { filter: "Estado eq 'abierto'" })
        const unicos = []
        const vistos = new Set()
        regs.forEach(r => {
          if (r.Codigo_Lote && !vistos.has(r.Codigo_Lote)) {
            vistos.add(r.Codigo_Lote)
            unicos.push({ lote: r.Codigo_Lote, producto: r.Producto, maquina: r.Title || '' })
          }
        })
        setLotesActivos(unicos)
      } catch { /* opcional */ }
    } catch (err) {
      console.error('Error cargando Kardex:', err)
    } finally {
      setCargando(false)
    }
  }, [])

  useEffect(() => { cargar() }, [cargar])

  // Cuando PCP selecciona máquina, cargar OFs activas para sugerir
  useEffect(() => {
    if (!maquina) { setOfsActivasMaq([]); return }
    let cancelled = false
    const cargarOFs = async () => {
      try {
        const token = await getToken()
        if (!token) return
        const ofs = await getOFsActivas(token, maquina)
        if (!cancelled) setOfsActivasMaq(ofs)
      } catch { setOfsActivasMaq([]) }
    }
    cargarOFs()
    return () => { cancelled = true }
  }, [maquina])

  const guardarAdicionales = async () => {
    const validas = filasAdicionales.filter(f => f.mp && parseFloat(f.kg) > 0)
    if (!validas.length || !modalAdicional) return
    setGuardandoAd(true)
    try {
      const token = await getToken()
      await Promise.all(validas.map(f => {
        const kgNuevo = parseFloat(f.kg)
        const insumoKey = (f.mp || '').trim().toLowerCase()
        // Buscar entrada PCP existente para este insumo en la misma OF
        // (no de operario — esas se mantienen separadas)
        const existente = entradas.find(e =>
          (e.Insumo || '').trim().toLowerCase() === insumoKey &&
          e.Numero_OF === modalAdicional.numeroOF &&
          !e.Observacion?.includes('operario')
        )
        if (existente) {
          // Actualizar: sumar kg a la entrada existente
          return updateListItem(token, 'Kardex_MP', existente.ID, {
            KgEntregados: (existente.KgEntregados || 0) + kgNuevo,
          })
        } else {
          // Crear nueva entrada solo si no existe para este insumo
          return createListItem(token, 'Kardex_MP', {
            Title: modalAdicional.maquina,
            Fecha: new Date(modalAdicional.fecha + 'T12:00:00').toISOString(),
            Turno: modalAdicional.turno,
            Insumo: f.mp,
            KgEntregados: kgNuevo,
            KgDevueltos: 0,
            Numero_OF: modalAdicional.numeroOF || '',
          })
        }
      }))
      setModalAdicional(null)
      setFilasAdicionales([filaVacia()])
      cargar()
    } catch (err) {
      alert('Error: ' + err.message)
    } finally {
      setGuardandoAd(false)
    }
  }

  const guardarEdicionEntrada = async (id) => {
    setGuardandoEntrada(true)
    try {
      const token = await getToken()
      const kg = parseFloat(editEntrada.KgEntregados) || 0
      await updateListItem(token, 'Kardex_MP', id, { KgEntregados: kg })
      setEntradas(prev => prev.map(e => e.ID === id ? { ...e, KgEntregados: kg } : e))
      setEditEntradaId(null)
    } catch (err) {
      alert('Error al guardar: ' + err.message)
    } finally {
      setGuardandoEntrada(false)
    }
  }

  // Validar: PCP confirma lo declarado por operario → KgEntregados = KgDeclaradoOperario
  // Luego recalcula KgMPRestante en Registro_Produccion del turno activo de esa OF
  const validarEntrada = async (entrada) => {
    try {
      const token = await getToken()
      const kgValidado = entrada.KgDeclaradoOperario ?? entrada.KgEntregados
      const sello = `Validado PCP · ${format(new Date(), 'dd/MM/yyyy HH:mm')}`
      const obsActual = (entrada.Observacion || '').replace('Registrada por operario', '').trim()
      const obsNueva = `${sello}${obsActual ? ' · ' + obsActual : ''}`
      await updateListItem(token, 'Kardex_MP', entrada.ID, {
        KgEntregados: kgValidado,
        Observacion: obsNueva,
      })

      // Recalcular KgMPRestante en el Registro_Produccion abierto de la OF
      if (entrada.Numero_OF) {
        const entradasActualizadas = entradas.map(e =>
          e.ID === entrada.ID ? { ...e, KgEntregados: kgValidado } : e
        )
        const kardexOF = entradasActualizadas.filter(e => e.Numero_OF === entrada.Numero_OF)
        const totalEntregado = kardexOF.reduce((s, e) => s + (e.KgEntregados || 0), 0)
        const totalUsado = kardexOF.reduce((s, e) =>
          s + (e.KgUsado || 0) + (e.KgMermaRec || 0) + (e.KgMermaNoRec || 0) + (e.KgDevueltos || 0), 0)
        const kgRestante = Math.max(0, totalEntregado - totalUsado)
        // Buscar el registro abierto de esta OF para actualizar KgMPRestante
        const regs = await getListItems(token, 'Registro_Produccion', { top: 200 })
        const regActivo = regs.find(r =>
          r.Numero_OF === entrada.Numero_OF && r.Estado !== 'cerrado'
        )
        if (regActivo) {
          await updateListItem(token, 'Registro_Produccion', regActivo.ID, {
            KgMPRestante: parseFloat(kgRestante.toFixed(3)),
          })
        }
      }

      setEntradas(prev => prev.map(e => e.ID === entrada.ID
        ? { ...e, KgEntregados: kgValidado, Observacion: obsNueva }
        : e
      ))
    } catch (err) {
      alert('Error al validar: ' + err.message)
    }
  }

  const yaValidado = (obs) => (obs || '').includes('Validado PCP')

  const guardarDevolucion = async (id) => {
    const kg = parseFloat(editDevKg)
    if (isNaN(kg) || kg < 0) return
    setGuardandoDev(true)
    try {
      const token = await getToken()
      await updateListItem(token, 'Kardex_MP', id, { KgDevueltos: kg })
      setEntradas(prev => prev.map(e => e.ID === id ? { ...e, KgDevueltos: kg } : e))
      setEditDevId(null)
      setEditDevKg('')
    } catch (err) {
      alert('Error al guardar devolución: ' + err.message)
    } finally {
      setGuardandoDev(false)
    }
  }

  // ── Submit ────────────────────────────────────────────────────────────────
  const handleSubmit = async (e) => {
    e.preventDefault()
    setFeedback(null)

    // Validaciones básicas
    if (!maquina) { setFeedback({ tipo: 'error', msg: 'Selecciona una máquina.' }); return }
    const filasValidas = filas.filter(f => f.mp && parseFloat(f.kg) > 0)
    if (filasValidas.length === 0) {
      setFeedback({ tipo: 'error', msg: 'Agrega al menos una materia prima con kg mayor a 0.' })
      return
    }

    setEnviando(true)
    try {
      const token = await getToken()
      // Crear un registro en SP por cada fila válida
      // Si PCP no seleccionó OF → generar automáticamente
      const ofFinal = codigoLote || `OF-${maquina}-${fecha.replace(/-/g,'')}-${format(new Date(), 'HHmm')}`

      await Promise.all(filasValidas.map(fila =>
        createListItem(token, 'Kardex_MP', {
          Title: maquina,
          Fecha: new Date(fecha).toISOString(),
          Turno: turno,
          Insumo: fila.mp,
          KgEntregados: parseFloat(fila.kg),
          KgDevueltos: 0,
          Observacion: obs || '',
          Numero_OF: ofFinal,
        })
      ))
      setCodigoLote(ofFinal) // mostrar la OF generada en el feedback

      const kgTotal = filasValidas.reduce((s, f) => s + parseFloat(f.kg), 0)
      setFeedback({
        tipo: 'exito',
        msg: `${filasValidas.length} registro(s) guardado(s) · Total: ${kgTotal.toFixed(2)} kg`,
      })
      // Reset filas pero mantener fecha/turno/máquina para agilizar múltiples envíos
      setFilas([filaVacia()])
      setObs('')
      cargar()
    } catch (err) {
      setFeedback({ tipo: 'error', msg: 'Error al guardar: ' + err.message })
    } finally {
      setEnviando(false)
    }
  }

  return (
    <div style={{ backgroundColor: '#f5f5f5', minHeight: '100vh' }}>
      {/* Header */}
      <header style={{ backgroundColor: '#37BEEC', color: 'white', padding: '12px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', position: 'sticky', top: 0, zIndex: 100 }}>
        <div>
          <p style={{ fontSize: '11px', opacity: 0.7, textTransform: 'uppercase' }}>Panel PCP</p>
          <h1 style={{ fontSize: '17px', fontWeight: 700 }}>Kardex de Materia Prima</h1>
        </div>
        <div style={{ display: 'flex', gap: '8px' }}>
          <button onClick={onVolver} style={{ backgroundColor: 'rgba(255,255,255,0.25)', color: 'white', border: '1.5px solid rgba(255,255,255,0.6)', borderRadius: '8px', padding: '8px 14px', fontSize: '13px', fontWeight: 600, cursor: 'pointer' }}>
            ← Validación
          </button>
          <button onClick={onLogout} style={{ backgroundColor: 'transparent', border: '1.5px solid rgba(255,255,255,0.5)', color: 'white', borderRadius: '8px', padding: '6px 12px', fontSize: '12px', cursor: 'pointer' }}>
            Salir
          </button>
        </div>
      </header>

      {/* Tabs */}
      <div style={{ backgroundColor: '#37BEEC', padding: '0 16px' }}>
        <div style={{ maxWidth: '720px', margin: '0 auto', display: 'flex', gap: '4px' }}>
          {[
            { id: 'entregas', label: '📋 Entregas registradas' },
            { id: 'registrar', label: '+ Registrar entrega' },
          ].map(t => (
            <button key={t.id} onClick={() => setTab(t.id)} style={{
              padding: '10px 18px', border: 'none', cursor: 'pointer',
              fontSize: '14px', fontWeight: 700,
              backgroundColor: tab === t.id ? 'white' : 'transparent',
              color: tab === t.id ? '#37BEEC' : 'rgba(255,255,255,0.75)',
              borderRadius: '8px 8px 0 0',
              borderBottom: tab === t.id ? '3px solid white' : '3px solid transparent',
            }}>{t.label}</button>
          ))}
        </div>
      </div>

      <div style={{ padding: '16px', maxWidth: '720px', margin: '0 auto', display: 'flex', flexDirection: 'column', gap: '20px' }}>

        {/* Formulario — tab Registrar */}
        {tab === 'registrar' && (
        <div style={{ backgroundColor: 'white', borderRadius: '14px', padding: '20px', boxShadow: '0 2px 8px rgba(0,0,0,0.08)' }}>
          <h2 style={{ color: '#1b5e20', fontSize: '16px', fontWeight: 700, marginBottom: '16px' }}>
            Registrar entrega de MP
          </h2>

          <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>

            {/* Fecha + Turno */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
              <div>
                <label style={lbl}>Fecha *</label>
                <input type="date" value={fecha} onChange={e => setFecha(e.target.value)} style={inp} />
              </div>
              <div>
                <label style={lbl}>Turno *</label>
                <div style={{ display: 'flex', gap: '6px' }}>
                  {TURNOS.map(t => (
                    <button key={t.id} type="button" onClick={() => setTurno(t.id)} style={{
                      flex: 1, padding: '10px 4px', borderRadius: '8px', border: '2px solid',
                      borderColor: turno === t.id ? '#1b5e20' : '#ddd',
                      backgroundColor: turno === t.id ? '#1b5e20' : 'white',
                      color: turno === t.id ? 'white' : '#333',
                      fontWeight: 600, fontSize: '13px', cursor: 'pointer',
                    }}>{t.label}</button>
                  ))}
                </div>
              </div>
            </div>

            {/* Máquina */}
            <div>
              <label style={lbl}>Máquina *</label>
              <select value={maquina} onChange={e => setMaquina(e.target.value)} style={sel}>
                <option value="">Seleccionar máquina...</option>
                {maquinas.map(m => {
                  const nombre = m.Nombre || m.Title || '(sin nombre)'
                  const codigo = m.Codigo ? ` (${m.Codigo})` : ''
                  return <option key={m.ID} value={m.Codigo || nombre}>{nombre}{codigo}</option>
                })}
              </select>
              {maquinas.length === 0 && !cargando && (
                <p style={{ fontSize: '12px', color: '#888', marginTop: '4px' }}>
                  Sin datos — verifica que la lista Maestro_Maquinas existe en SharePoint.
                </p>
              )}
            </div>

            {/* OF eliminada del formulario — se asigna automáticamente o desde Entregas */}

            {/* Filas dinámicas de MP */}
            <div>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                <label style={lbl}>Materias primas entregadas *</label>
                <button type="button" onClick={agregarFila} style={{
                  backgroundColor: '#1b5e20', color: 'white', border: 'none',
                  borderRadius: '8px', padding: '6px 14px', fontSize: '13px',
                  fontWeight: 700, cursor: 'pointer',
                }}>+ Agregar MP</button>
              </div>

              {filas.map((fila) => (
                <div key={fila.id} style={{ display: 'grid', gridTemplateColumns: '1fr 80px 32px', gap: '6px', alignItems: 'end', marginBottom: '8px' }}>
                  <div>
                    <label style={{ ...lbl, fontSize: '11px' }}>Insumo</label>
                    <SearchSelect
                      opciones={materiasPrimas.map(mp => ({
                        value: mp.Nombre || mp.Title || '',
                        label: (mp.Nombre || mp.Title || '') + (mp.Codigo ? ` (${mp.Codigo})` : ''),
                      }))}
                      value={fila.mp}
                      onChange={v => updateFila(fila.id, 'mp', v)}
                      placeholder="Buscar insumo…"
                    />
                  </div>
                  <div>
                    <label style={{ ...lbl, fontSize: '11px' }}>Kg</label>
                    <input type="number" step="0.01" min="0"
                      value={fila.kg || ''}
                      onChange={e => updateFila(fila.id, 'kg', e.target.value)}
                      placeholder="0.00"
                      style={{ ...inp, textAlign: 'right', fontWeight: 700, fontSize: '16px', padding: '14px 8px', boxSizing: 'border-box' }} />
                  </div>
                  <button type="button" onClick={() => eliminarFila(fila.id)}
                    disabled={filas.length === 1}
                    style={{
                      backgroundColor: filas.length === 1 ? 'transparent' : '#ffebee',
                      color: filas.length === 1 ? '#ccc' : '#c62828',
                      border: 'none', borderRadius: '6px',
                      width: '32px', height: '52px', fontSize: '16px',
                      cursor: filas.length === 1 ? 'not-allowed' : 'pointer',
                    }}>×</button>
                </div>
              ))}

              {filas.some(f => parseFloat(f.kg) > 0) && (
                <div style={{ display: 'flex', justifyContent: 'flex-end', paddingTop: '8px', borderTop: '2px solid #e0e0e0', marginTop: '4px' }}>
                  <span style={{ fontSize: '16px', fontWeight: 800, color: '#1b5e20' }}>
                    Total: {filas.reduce((s, f) => s + (parseFloat(f.kg) || 0), 0).toFixed(2)} kg
                  </span>
                </div>
              )}
            </div>

            {/* Observaciones */}
            <div>
              <label style={lbl}>Observaciones</label>
              <textarea value={obs} onChange={e => setObs(e.target.value)} rows={2}
                placeholder="Notas adicionales sobre la entrega..."
                style={{ ...inp, resize: 'none', fontFamily: 'inherit' }} />
            </div>

            {feedback && (
              <div style={{
                padding: '12px', borderRadius: '8px',
                backgroundColor: feedback.tipo === 'exito' ? '#e8f5e9' : '#ffebee',
                color: feedback.tipo === 'exito' ? '#2e7d32' : '#c62828',
                fontSize: '14px', fontWeight: 600,
              }}>
                {feedback.tipo === 'exito' ? '✓ ' : '✗ '}{feedback.msg}
              </div>
            )}

            <button type="submit" disabled={enviando} style={{
              backgroundColor: enviando ? '#ccc' : '#1b5e20', color: 'white',
              border: 'none', borderRadius: '12px', padding: '16px',
              fontSize: '16px', fontWeight: 700, minHeight: '56px',
              cursor: enviando ? 'not-allowed' : 'pointer',
            }}>
              {enviando
                ? '⏳ Guardando...'
                : `✓ Registrar ${filas.filter(f => f.mp && parseFloat(f.kg) > 0).length || ''} entrega(s) en Kardex`}
            </button>
          </form>
        </div>
        )} {/* fin tab registrar */}

        {/* Tab Entregas registradas */}
        {tab === 'entregas' && (
        <div style={{ backgroundColor: 'white', borderRadius: '14px', padding: '20px', boxShadow: '0 2px 8px rgba(0,0,0,0.08)' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '12px' }}>
            <input type="date" value={filtroFechaHist} onChange={e => setFiltroFechaHist(e.target.value)}
              style={{ padding: '7px 8px', borderRadius: '8px', border: '2px solid #ddd', fontSize: '13px', color: '#1a1a1a', backgroundColor: 'white', flex: '1 1 0', minWidth: 0 }} />
            <select value={filtroTurnoHist} onChange={e => setFiltroTurnoHist(e.target.value)}
              style={{ padding: '7px 6px', borderRadius: '8px', border: '2px solid #ddd', fontSize: '13px', color: '#1a1a1a', backgroundColor: 'white', width: '90px', flexShrink: 0 }}>
              <option value="todos">Todos</option>
              <option value="M">Mañana</option>
              <option value="T">Tarde</option>
              <option value="N">Noche</option>
            </select>
            <button onClick={() => { setFiltroFechaHist(''); setFiltroTurnoHist('todos') }}
              style={{ padding: '7px 10px', borderRadius: '8px', border: '1px solid #ddd', backgroundColor: 'white', color: '#555', fontSize: '12px', cursor: 'pointer', whiteSpace: 'nowrap', flexShrink: 0 }}>
              Ver todo
            </button>
          </div>
          {cargando ? (
            <p style={{ color: '#888', fontSize: '14px' }}>Cargando...</p>
          ) : (() => {
            // Filtrar por fecha
            const filtradas = entradas.slice(0, 200).filter(e => {
              const matchFecha = !filtroFechaHist || (e.Fecha || '').split('T')[0] === filtroFechaHist
              const matchTurno = filtroTurnoHist === 'todos' || e.Turno === filtroTurnoHist
              return matchFecha && matchTurno
            })

            // Agrupar por Maquina + OF + Fecha + Turno
            const grupos = {}
            filtradas.forEach(e => {
              const f    = (e.Fecha || '').split('T')[0]
              const of   = e.Numero_OF || 'sin-OF'
              const key  = `${e.Title || ''}|${of}|${f}|${e.Turno || ''}`
              if (!grupos[key]) grupos[key] = {
                key, maquina: e.Title || '—', fecha: f, turno: e.Turno,
                numeroOF: e.Numero_OF || null, items: [],
              }
              grupos[key].items.push(e)
            })
            const lista = Object.values(grupos)

            if (!lista.length) return (
              <p style={{ color: '#aaa', fontSize: '14px', fontStyle: 'italic' }}>
                Sin entregas para la fecha seleccionada.
              </p>
            )

            return (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                {lista.map(grupo => {
                  const totalEnt = grupo.items.reduce((s, i) => s + (i.KgEntregados || 0), 0)
                  const totalDev = grupo.items.reduce((s, i) => s + (i.KgDevueltos || 0), 0)
                  const neto = totalEnt - totalDev
                  const abierto = gruposAbiertos[grupo.key] === true

                  // Estado del grupo para el badge esquina
                  const tieneOpSinValidar = grupo.items.some(i => {
                    // Caso 1: entrada marcada explícitamente por operario
                    if (i.Observacion?.includes('operario') && !yaValidado(i.Observacion)) return true
                    // Caso 2: KgDeclaradoOperario difiere de KgEntregados (operario declaró diferente)
                    if (i.KgDeclaradoOperario != null &&
                        Math.abs((i.KgDeclaradoOperario || 0) - (i.KgEntregados || 0)) > 0.01 &&
                        !yaValidado(i.Observacion)) return true
                    return false
                  })
                  const todosValidados = grupo.items.length > 0 &&
                    grupo.items.every(i => yaValidado(i.Observacion))
                  const grupoEstado = todosValidados ? 'validado'
                    : tieneOpSinValidar ? 'operario' : null

                  // Colores según estado
                  const bgHeader = tieneOpSinValidar ? '#e65100' : '#1b5e20'
                  const borderColor = tieneOpSinValidar ? '#ff9800' : '#c8e6c9'

                  return (
                    <div key={grupo.key} style={{
                      borderRadius: '12px',
                      border: `1.5px solid ${borderColor}`,
                      overflow: 'hidden',
                      boxShadow: '0 2px 6px rgba(0,0,0,0.08)',
                    }}>
                      {/* Cabecera */}
                      <div
                        onClick={() => toggleGrupo(grupo.key)}
                        style={{ backgroundColor: bgHeader, color: 'white', padding: '12px 14px', cursor: 'pointer' }}
                      >
                        {/* Fila 1: máquina + kg + estado + flecha */}
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                            <strong style={{ fontSize: '16px' }}>{grupo.maquina}</strong>
                            <span style={{ fontSize: '12px', opacity: 0.8, backgroundColor: 'rgba(255,255,255,0.15)', borderRadius: '6px', padding: '1px 8px' }}>
                              Turno {grupo.turno}
                            </span>
                            <span style={{ fontSize: '12px', opacity: 0.75 }}>
                              {grupo.fecha ? format(new Date(grupo.fecha + 'T12:00:00'), 'dd/MM/yyyy') : '—'}
                            </span>
                          </div>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                            <strong style={{ fontSize: '18px' }}>{neto.toFixed(2)} kg</strong>
                            <span style={{ opacity: 0.6, fontSize: '16px' }}>{abierto ? '▲' : '▼'}</span>
                          </div>
                        </div>

                        {/* Fila 2: OF + desglose + badge */}
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '6px' }}>
                          <div style={{ fontSize: '11px', opacity: 0.75, display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                            {grupo.numeroOF && (
                              <span style={{ fontFamily: 'monospace' }}>🔖 {grupo.numeroOF}</span>
                            )}
                            {totalDev > 0 && (
                              <span>{totalEnt} ent. − {totalDev} dev.</span>
                            )}
                          </div>
                          {/* Badge estado */}
                          {grupoEstado === 'validado' && (
                            <span style={{
                              backgroundColor: 'rgba(255,255,255,0.2)',
                              border: '1px solid rgba(255,255,255,0.5)',
                              borderRadius: '10px', padding: '2px 10px',
                              fontSize: '11px', fontWeight: 700,
                            }}>✓ Validado PCP</span>
                          )}
                          {tieneOpSinValidar && (
                            <span style={{
                              backgroundColor: 'rgba(255,255,255,0.2)',
                              border: '1px solid rgba(255,255,255,0.5)',
                              borderRadius: '10px', padding: '2px 10px',
                              fontSize: '11px', fontWeight: 700,
                            }}>👤 Por operario · pendiente validación</span>
                          )}
                        </div>
                      </div>

                      {/* Items del grupo — consolidados por insumo */}
                      {abierto && (
                        <div style={{ backgroundColor: 'white' }}>
                          {(() => {
                            const porInsumo = {}
                            grupo.items.forEach(e => {
                              const key = (e.Insumo || '—').trim().toLowerCase()
                              if (!porInsumo[key]) porInsumo[key] = []
                              porInsumo[key].push(e)
                            })
                            return Object.entries(porInsumo).map(([, entries], gIdx) => {
                              const nombre = entries[0].Insumo || '—'
                              // Total: si hay declaración pendiente del op, usar KgDeclaradoOperario como target final
                              const totalKg = entries.reduce((s, e) => {
                                const pendiente = e.Observacion?.includes('operario') && !yaValidado(e.Observacion)
                                const kg = pendiente ? (e.KgDeclaradoOperario || e.KgEntregados || 0) : (e.KgEntregados || 0)
                                return s + kg
                              }, 0)
                              const totalDev = entries.reduce((s, e) => s + (e.KgDevueltos || 0), 0)
                              const tieneAdicional = entries.length > 1
                              const tieneOp = entries.some(e => e.Observacion?.includes('operario'))
                              return (
                                <div key={nombre} style={{ padding: '12px 16px', borderTop: gIdx > 0 ? '1px solid #f0f0f0' : 'none' }}>

                                  {/* Fila principal: nombre + total */}
                                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: entries.length > 0 ? '8px' : 0 }}>
                                    <div>
                                      <span style={{ fontSize: '14px', fontWeight: 600, color: '#1a1a1a' }}>{nombre}</span>
                                      {tieneOp && (
                                        <span style={{ marginLeft: '8px', backgroundColor: '#fff3e0', color: '#e65100', borderRadius: '4px', padding: '1px 6px', fontSize: '10px', fontWeight: 700 }}>OP</span>
                                      )}
                                    </div>
                                    <span style={{ fontSize: '16px', fontWeight: 700, color: '#1b5e20' }}>
                                      {totalKg.toFixed(2)} kg
                                      {totalDev > 0 && <span style={{ fontSize: '12px', color: '#f57f17', marginLeft: '6px' }}>−{totalDev.toFixed(2)}</span>}
                                    </span>
                                  </div>

                                  {/* Entradas individuales — compactas */}
                                  {entries.map((e, eIdx) => {
                                    const hayDiscrepancia = e.KgDeclaradoOperario != null &&
                                      Math.abs((e.KgDeclaradoOperario || 0) - (e.KgEntregados || 0)) > 0.01
                                    const esOp = e.Observacion?.includes('operario') || hayDiscrepancia
                                    const esPcpAd = e.Observacion === 'MP adicional'
                                    const pendienteOp = esOp && !yaValidado(e.Observacion)
                                    // Delta = lo nuevo que el operario declaró (aún no en KgEntregados)
                                    const delta = pendienteOp
                                      ? (e.KgDeclaradoOperario || 0) - (e.KgEntregados || 0)
                                      : 0
                                    const kgEntrada = pendienteOp
                                      ? (e.KgDeclaradoOperario || 0)
                                      : (e.KgEntregados || e.KgDeclaradoOperario || 0)
                                    const badgeColor = esOp ? { bg: '#fff3e0', text: '#e65100' } : esPcpAd ? { bg: '#e3f2fd', text: '#1565c0' } : { bg: '#f3f4f6', text: '#555' }
                                    const badgeLabel = esOp ? 'Operario' : esPcpAd ? 'PCP adicional' : 'PCP'
                                    return (
                                      <div key={e.ID} style={{ marginTop: eIdx > 0 ? '6px' : 0, borderLeft: `3px solid ${badgeColor.bg === '#f3f4f6' ? '#e0e0e0' : badgeColor.bg}`, paddingLeft: '10px' }}>
                                        {/* Fila de entrada */}
                                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                            <span style={{ backgroundColor: badgeColor.bg, color: badgeColor.text, borderRadius: '4px', padding: '1px 7px', fontSize: '10px', fontWeight: 700 }}>{badgeLabel}</span>
                                            <span style={{ fontSize: '13px', color: '#333' }}>{kgEntrada} kg</span>
                                            {pendienteOp && delta > 0 && (
                                              <span style={{ fontSize: '11px', fontWeight: 700, color: '#e65100', backgroundColor: '#fff3e0', borderRadius: '4px', padding: '1px 6px' }}>
                                                +{delta.toFixed(2)} kg a validar
                                              </span>
                                            )}
                                            {e.KgDevueltos > 0 && <span style={{ fontSize: '11px', color: '#f57f17' }}>↩ {e.KgDevueltos} dev.</span>}
                                            {esOp && !yaValidado(e.Observacion) && (
                                              <button type="button" onClick={() => validarEntrada(e)}
                                                style={{ backgroundColor: '#e8f5e9', color: '#1b5e20', border: '1px solid #a5d6a7', borderRadius: '4px', padding: '1px 8px', fontSize: '10px', fontWeight: 700, cursor: 'pointer' }}>✓ Validar</button>
                                            )}
                                          </div>
                                          <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
                                            <button type="button" onClick={() => { setEditEntradaId(editEntradaId === e.ID ? null : e.ID); setEditEntrada({ KgEntregados: e.KgEntregados }) }}
                                              style={{ backgroundColor: '#f5f5f5', border: '1px solid #e0e0e0', color: '#555', borderRadius: '4px', padding: '1px 7px', fontSize: '10px', fontWeight: 600, cursor: 'pointer' }}>✏ Editar</button>
                                            <button type="button" onClick={() => { setEditDevId(editDevId === e.ID ? null : e.ID); setEditDevKg(String(e.KgDevueltos || 0)) }}
                                              style={{ backgroundColor: '#fff8f0', border: '1px solid #ffcc80', color: '#f57f17', borderRadius: '4px', padding: '1px 7px', fontSize: '10px', fontWeight: 600, cursor: 'pointer' }}>↩ Dev.</button>
                                          </div>
                                        </div>
                                        {/* Edición inline */}
                                        {editEntradaId === e.ID && (
                                          <div style={{ display: 'flex', gap: '6px', alignItems: 'center', marginTop: '6px' }}>
                                            <span style={{ fontSize: '11px', color: '#555', whiteSpace: 'nowrap' }}>Kg entregados:</span>
                                            <input type="number" step="0.01" min="0" value={editEntrada.KgEntregados ?? ''} onChange={ev => setEditEntrada(p => ({ ...p, KgEntregados: ev.target.value }))} autoFocus style={{ ...inp, flex: 1, padding: '5px 8px', textAlign: 'right', fontWeight: 700, fontSize: '13px' }} />
                                            <button onClick={() => guardarEdicionEntrada(e.ID)} disabled={guardandoEntrada} style={{ backgroundColor: '#1b5e20', color: 'white', border: 'none', borderRadius: '6px', padding: '5px 12px', fontSize: '12px', fontWeight: 700, cursor: 'pointer' }}>{guardandoEntrada ? '...' : '✓'}</button>
                                            <button onClick={() => setEditEntradaId(null)} style={{ background: 'none', border: '1px solid #ddd', borderRadius: '6px', padding: '5px 10px', fontSize: '12px', color: '#666', cursor: 'pointer' }}>✕</button>
                                          </div>
                                        )}
                                        {/* Devolución inline */}
                                        {editDevId === e.ID && (
                                          <div style={{ display: 'flex', gap: '6px', alignItems: 'center', marginTop: '6px' }}>
                                            <span style={{ fontSize: '11px', color: '#555', whiteSpace: 'nowrap' }}>Kg devueltos:</span>
                                            <input type="number" step="0.01" min="0" value={editDevKg} onChange={ev => setEditDevKg(ev.target.value)} autoFocus style={{ ...inp, flex: 1, padding: '5px 8px', textAlign: 'right', fontWeight: 700, fontSize: '13px' }} />
                                            <button onClick={() => guardarDevolucion(e.ID)} disabled={guardandoDev} style={{ backgroundColor: '#1b5e20', color: 'white', border: 'none', borderRadius: '6px', padding: '5px 12px', fontSize: '12px', fontWeight: 700, cursor: 'pointer' }}>{guardandoDev ? '...' : '✓'}</button>
                                            <button onClick={() => setEditDevId(null)} style={{ background: 'none', border: '1px solid #ddd', borderRadius: '6px', padding: '5px 10px', fontSize: '12px', color: '#666', cursor: 'pointer' }}>✕</button>
                                          </div>
                                        )}
                                      </div>
                                    )
                                  })}
                                </div>
                              )
                            })
                          })()}
                          {/* Botón MP adicional al final */}
                          <div style={{ padding: '10px 14px', borderTop: '1px solid #f0f0f0', display: 'flex', justifyContent: 'flex-end' }}>
                            <button
                              onClick={() => {
                                setModalAdicional({ maquina: grupo.maquina, turno: grupo.turno, fecha: grupo.fecha, numeroOF: grupo.numeroOF })
                                setFilasAdicionales([filaVacia()])
                              }}
                              style={{
                                backgroundColor: 'transparent', color: '#1b5e20',
                                border: '1.5px solid #1b5e20', borderRadius: '8px',
                                padding: '6px 16px', fontSize: '12px', fontWeight: 700, cursor: 'pointer',
                              }}>
                              + MP adicional
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  )
                })}
              </div>
            )
          })()}
        </div>
        )} {/* fin tab entregas */}

      </div>

      {/* Modal MP adicional a grupo existente */}
      {modalAdicional && (
        <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 300, padding: '16px' }}>
          <div style={{ backgroundColor: 'white', borderRadius: '16px', padding: '24px', width: '100%', maxWidth: '460px', display: 'flex', flexDirection: 'column', gap: '16px', maxHeight: '90vh', overflowY: 'auto' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <h3 style={{ color: '#1b5e20', fontSize: '16px', fontWeight: 700 }}>Agregar MP adicional</h3>
                <p style={{ fontSize: '13px', color: '#555', marginTop: '2px' }}>
                  {modalAdicional.maquina} · Turno {modalAdicional.turno} · {format(new Date(modalAdicional.fecha + 'T12:00:00'), 'dd/MM/yyyy')}
                </p>
              </div>
              <button onClick={() => setModalAdicional(null)} style={{ background: 'none', border: 'none', fontSize: '22px', cursor: 'pointer', color: '#666' }}>✕</button>
            </div>

            {filasAdicionales.map(fila => (
              <div key={fila.id} style={{ display: 'grid', gridTemplateColumns: '1fr 80px 32px', gap: '6px', alignItems: 'end' }}>
                <div>
                  <label style={{ ...lbl, fontSize: '11px' }}>Insumo</label>
                  <SearchSelect
                    opciones={materiasPrimas.map(mp => ({
                      value: mp.Nombre || mp.Title || '',
                      label: (mp.Nombre || mp.Title || '') + (mp.Codigo ? ` (${mp.Codigo})` : ''),
                    }))}
                    value={fila.mp}
                    onChange={v => setFilasAdicionales(p => p.map(f => f.id === fila.id ? { ...f, mp: v } : f))}
                    placeholder="Buscar insumo…"
                  />
                </div>
                <div>
                  <label style={{ ...lbl, fontSize: '11px' }}>Kg</label>
                  <input type="number" step="0.01" min="0" value={fila.kg}
                    onChange={e => setFilasAdicionales(p => p.map(f => f.id === fila.id ? { ...f, kg: e.target.value } : f))}
                    placeholder="0.00"
                    style={{ ...inp, textAlign: 'right', fontWeight: 700, padding: '14px 8px', boxSizing: 'border-box' }} />
                </div>
                <button type="button" onClick={() => setFilasAdicionales(p => p.length > 1 ? p.filter(f => f.id !== fila.id) : p)}
                  style={{ backgroundColor: filasAdicionales.length === 1 ? '#f5f5f5' : '#ffebee', color: filasAdicionales.length === 1 ? '#ccc' : '#c62828', border: 'none', borderRadius: '6px', width: '32px', height: '52px', fontSize: '16px', cursor: 'pointer' }}>×</button>
              </div>
            ))}

            <button type="button" onClick={() => setFilasAdicionales(p => [...p, filaVacia()])} style={{
              backgroundColor: '#e8f5e9', color: '#1b5e20', border: '1px solid #1b5e20',
              borderRadius: '8px', padding: '8px', fontSize: '13px', fontWeight: 700, cursor: 'pointer',
            }}>+ Agregar otra MP</button>

            {filasAdicionales.some(f => parseFloat(f.kg) > 0) && (
              <p style={{ fontSize: '13px', fontWeight: 700, color: '#1b5e20', textAlign: 'right' }}>
                Total: {filasAdicionales.reduce((s, f) => s + (parseFloat(f.kg) || 0), 0).toFixed(2)} kg
              </p>
            )}

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', gap: '10px' }}>
              <button onClick={() => setModalAdicional(null)} style={{ padding: '12px', borderRadius: '10px', border: '2px solid #ccc', backgroundColor: 'white', color: '#555', fontSize: '14px', cursor: 'pointer' }}>Cancelar</button>
              <button onClick={guardarAdicionales} disabled={guardandoAd} style={{
                padding: '12px', borderRadius: '10px', border: 'none',
                backgroundColor: guardandoAd ? '#ccc' : '#1b5e20', color: 'white',
                fontSize: '14px', fontWeight: 700, cursor: guardandoAd ? 'not-allowed' : 'pointer',
              }}>{guardandoAd ? '⏳ Guardando...' : '✓ Guardar MP adicional'}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
