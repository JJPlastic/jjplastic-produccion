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

const filaVacia  = () => ({ id: crypto.randomUUID(), mp: '', kg: '' })
const grupoVacio = () => ({ id: crypto.randomUUID(), producto: '', filas: [filaVacia()] })

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
  const [codigoLote, setCodigoLote] = useState('')
  const [lotesActivos, setLotesActivos] = useState([])
  const [ofsActivasMaq, setOfsActivasMaq] = useState([]) // OFs activas cuando se selecciona máquina

  // ── Grupos de entrega (producto → sus MPs) ───────────────────────────────
  const [grupos, setGrupos] = useState([grupoVacio()])

  const agregarGrupo = () => setGrupos(prev => [...prev, grupoVacio()])
  const eliminarGrupo = (gid) => setGrupos(prev => prev.length > 1 ? prev.filter(g => g.id !== gid) : prev)
  const updateGrupoProducto = (gid, val) =>
    setGrupos(prev => prev.map(g => g.id === gid ? { ...g, producto: val } : g))
  const agregarFilaGrupo = (gid) =>
    setGrupos(prev => prev.map(g => g.id === gid && g.filas.length < 5 ? { ...g, filas: [...g.filas, filaVacia()] } : g))
  const eliminarFilaGrupo = (gid, fid) =>
    setGrupos(prev => prev.map(g => g.id === gid ? { ...g, filas: g.filas.length > 1 ? g.filas.filter(f => f.id !== fid) : g.filas } : g))
  const updateFilaGrupo = (gid, fid, campo, valor) =>
    setGrupos(prev => prev.map(g => g.id === gid ? { ...g, filas: g.filas.map(f => f.id === fid ? { ...f, [campo]: valor } : f) } : g))

  // ── Datos ─────────────────────────────────────────────────────────────────
  const [maquinas, setMaquinas]             = useState([])
  const [materiasPrimas, setMateriasPrimas] = useState([])
  const [productosPA, setProductosPA]       = useState([])
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
  // OFs con transferencia pendiente (flag desde Registro_Produccion)
  const [ofsConTransferencia, setOfsConTransferencia] = useState(new Set())
  // OFs que ya tienen al menos un Registro_Produccion (operario inició producción)
  const [ofsConRegistro, setOfsConRegistro] = useState(new Set())
  // Modales reasignar / transferir
  const [modalReasignar, setModalReasignar]   = useState(null) // { grupo }
  const [modalTransferir, setModalTransferir] = useState(null) // { grupo, kgPorInsumo }
  const [maqDestinoSel, setMaqDestinoSel]     = useState('')
  const [guardandoReasig, setGuardandoReasig] = useState(false)
  const [guardandoTransf, setGuardandoTransf] = useState(false)

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
        const todos = prodsRes.value.filter(esActivo)
        setMateriasPrimas(todos.filter(p =>
          ['MP', 'MC'].includes((p.TipoProducto || '').toUpperCase())
        ))
        setProductosPA(todos.filter(p =>
          !['MP', 'MC'].includes((p.TipoProducto || '').toUpperCase())
        ))
      }
      if (kardexRes.status === 'fulfilled') setEntradas(kardexRes.value)

      // OFs con Transferencia_Pendiente=true y OFs con registro de producción
      try {
        const regsTransf = await getListItems(token, 'Registro_Produccion', { top: 200 })
        const conFlag = new Set(
          regsTransf
            .filter(r => r.Transferencia_Pendiente === true)
            .map(r => r.Numero_OF)
            .filter(Boolean)
        )
        setOfsConTransferencia(conFlag)
        const conRegistro = new Set(
          regsTransf.map(r => r.Numero_OF).filter(Boolean)
        )
        setOfsConRegistro(conRegistro)
      } catch { /* opcional */ }

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

  // ── Reasignar máquina (error de asignación) ───────────────────────────────
  const ejecutarReasignar = async () => {
    if (!maqDestinoSel || !modalReasignar) return
    setGuardandoReasig(true)
    try {
      const token = await getToken()
      const { grupo } = modalReasignar
      const nota = `Reasignado desde ${grupo.maquina} · ${format(new Date(), 'dd/MM/yyyy HH:mm')}`
      await Promise.all(grupo.items.map(e =>
        updateListItem(token, 'Kardex_MP', e.ID, {
          Title: maqDestinoSel,
          Observacion: e.Observacion ? `${nota} · ${e.Observacion}` : nota,
        })
      ))
      setModalReasignar(null)
      setMaqDestinoSel('')
      cargar()
    } catch (err) {
      alert('Error al reasignar: ' + err.message)
    } finally {
      setGuardandoReasig(false)
    }
  }

  // ── Transferir kg restantes a otra máquina (máquina malograda) ────────────
  const ejecutarTransferir = async () => {
    if (!maqDestinoSel || !modalTransferir) return
    setGuardandoTransf(true)
    try {
      const token = await getToken()
      const { grupo } = modalTransferir
      const ahora = new Date()
      const nota = `Transferido a ${maqDestinoSel} · ${format(ahora, 'dd/MM/yyyy HH:mm')}`

      // Calcular kg disponibles por insumo (no usados/devueltos aún)
      const porInsumo = {}
      grupo.items.forEach(e => {
        const key = (e.Insumo || '').trim().toLowerCase()
        if (!porInsumo[key]) porInsumo[key] = { nombre: e.Insumo, disponible: 0, entries: [] }
        const disp = Math.max(0,
          (e.KgEntregados || 0) - (e.KgUsado || 0) - (e.KgMermaRec || 0)
          - (e.KgMermaNoRec || 0) - (e.KgDevueltos || 0)
        )
        porInsumo[key].disponible += disp
        porInsumo[key].entries.push({ ...e, disp })
      })

      const insumosATransferir = Object.values(porInsumo).filter(i => i.disponible > 0.001)
      if (!insumosATransferir.length) {
        alert('No hay kg disponibles para transferir en este grupo.')
        setGuardandoTransf(false)
        return
      }

      // Crear entradas nuevas en máquina destino
      await Promise.all(insumosATransferir.map(insumo =>
        createListItem(token, 'Kardex_MP', {
          Title: maqDestinoSel,
          Fecha: ahora.toISOString(),
          Turno: grupo.turno,
          Insumo: insumo.nombre,
          KgEntregados: parseFloat(insumo.disponible.toFixed(3)),
          KgDevueltos: 0,
          Numero_OF: grupo.numeroOF || '',
          Observacion: `Recibido desde ${grupo.maquina} · ${format(ahora, 'dd/MM/yyyy HH:mm')}`,
        })
      ))

      // Marcar kg como transferidos en entradas originales (suman a KgDevueltos con nota)
      await Promise.all(
        insumosATransferir.flatMap(insumo =>
          insumo.entries.filter(e => e.disp > 0.001).map(e =>
            updateListItem(token, 'Kardex_MP', e.ID, {
              KgDevueltos: parseFloat(((e.KgDevueltos || 0) + e.disp).toFixed(3)),
              Observacion: e.Observacion ? `${nota} · ${e.Observacion}` : nota,
            })
          )
        )
      )

      // Limpiar flag Transferencia_Pendiente en Registro_Produccion si estaba activo
      if (grupo.numeroOF && ofsConTransferencia.has(grupo.numeroOF)) {
        try {
          const regs = await getListItems(token, 'Registro_Produccion', { top: 200 })
          const conFlag = regs.filter(r =>
            r.Numero_OF === grupo.numeroOF && r.Transferencia_Pendiente === true
          )
          await Promise.all(conFlag.map(r =>
            updateListItem(token, 'Registro_Produccion', r.ID, { Transferencia_Pendiente: false })
          ))
        } catch { /* no crítico */ }
      }

      setModalTransferir(null)
      setMaqDestinoSel('')
      cargar()
    } catch (err) {
      alert('Error al transferir: ' + err.message)
    } finally {
      setGuardandoTransf(false)
    }
  }

  // ── Submit ────────────────────────────────────────────────────────────────
  const handleSubmit = async (e) => {
    e.preventDefault()
    setFeedback(null)

    if (!maquina) { setFeedback({ tipo: 'error', msg: 'Selecciona una máquina.' }); return }

    // Aplanar grupos → items válidos (producto + mp + kg)
    const items = grupos.flatMap(g => {
      const filasValidas = g.filas.filter(f => f.mp && parseFloat(f.kg) > 0)
      return filasValidas.map(f => ({ producto: g.producto, mp: f.mp, kg: parseFloat(f.kg) }))
    })

    if (items.length === 0) {
      setFeedback({ tipo: 'error', msg: 'Agrega al menos un insumo con kg mayor a 0.' })
      return
    }
    const gruposSinProducto = grupos.some(g =>
      g.filas.some(f => f.mp && parseFloat(f.kg) > 0) && !g.producto
    )
    if (gruposSinProducto) {
      setFeedback({ tipo: 'error', msg: 'Cada grupo de insumos debe tener un producto asignado.' })
      return
    }

    setEnviando(true)
    try {
      const token = await getToken()
      const ofFinal = codigoLote || `OF-${maquina}-${fecha.replace(/-/g,'')}-${format(new Date(), 'HHmm')}`

      await Promise.all(items.map(item =>
        createListItem(token, 'Kardex_MP', {
          Title: maquina,
          Fecha: new Date(fecha).toISOString(),
          Turno: turno,
          Insumo: item.mp,
          KgEntregados: item.kg,
          KgDevueltos: 0,
          Observacion: obs || '',
          Numero_OF: ofFinal,
          Producto: item.producto || '',
        })
      ))
      setCodigoLote(ofFinal)

      const kgTotal = items.reduce((s, i) => s + i.kg, 0)
      const nProductos = new Set(items.map(i => i.producto).filter(Boolean)).size
      setFeedback({
        tipo: 'exito',
        msg: `${items.length} insumo(s) registrados · ${nProductos} producto(s) · Total: ${kgTotal.toFixed(2)} kg`,
      })
      setGrupos([grupoVacio()])
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
        <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>

          {/* Sección 1: Fecha / Turno / Máquina */}
          <div style={{ backgroundColor: 'white', borderRadius: '14px', overflow: 'hidden', boxShadow: '0 2px 8px rgba(0,0,0,0.08)' }}>
            <div style={{ backgroundColor: '#004895', padding: '10px 16px', display: 'flex', alignItems: 'center', gap: '10px' }}>
              <span style={{ backgroundColor: 'rgba(255,255,255,0.2)', borderRadius: '50%', width: '22px', height: '22px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '12px', fontWeight: 800, color: 'white', flexShrink: 0 }}>1</span>
              <p style={{ fontSize: '11px', color: 'rgba(255,255,255,0.9)', textTransform: 'uppercase', letterSpacing: '0.06em', margin: 0, fontWeight: 700 }}>Programación</p>
            </div>
            <div style={{ padding: '14px 16px', display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                <div>
                  <label style={lbl}>Fecha *</label>
                  <input type="date" value={fecha} onChange={e => setFecha(e.target.value)} style={inp} />
                </div>
                <div>
                  <label style={lbl}>Turno *</label>
                  <div style={{ display: 'flex', gap: '4px' }}>
                    {TURNOS.map(t => (
                      <button key={t.id} type="button" onClick={() => setTurno(t.id)} style={{
                        flex: 1, padding: '11px 4px', borderRadius: '8px', border: '2px solid',
                        borderColor: turno === t.id ? '#004895' : '#ddd',
                        backgroundColor: turno === t.id ? '#004895' : 'white',
                        color: turno === t.id ? 'white' : '#555',
                        fontWeight: 700, fontSize: '13px', cursor: 'pointer',
                      }}>{t.label}</button>
                    ))}
                  </div>
                </div>
              </div>
              <div>
                <label style={lbl}>Máquina *</label>
                <select value={maquina} onChange={e => setMaquina(e.target.value)} style={{ ...sel, borderColor: maquina ? '#004895' : '#ddd', borderWidth: '2px' }}>
                  <option value="">Seleccionar máquina...</option>
                  {maquinas.map(m => {
                    const nombre = m.Nombre || m.Title || '(sin nombre)'
                    const codigo = m.Codigo ? ` (${m.Codigo})` : ''
                    return <option key={m.ID} value={m.Codigo || nombre}>{nombre}{codigo}</option>
                  })}
                </select>
              </div>
            </div>
          </div>

          {/* Sección 2: Grupos producto → MPs */}
          <div style={{ backgroundColor: 'white', borderRadius: '14px', boxShadow: '0 2px 8px rgba(0,0,0,0.08)' }}>
            <div style={{ backgroundColor: '#004895', padding: '10px 16px', borderRadius: '14px 14px 0 0', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <span style={{ backgroundColor: 'rgba(255,255,255,0.2)', borderRadius: '50%', width: '22px', height: '22px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '12px', fontWeight: 800, color: 'white', flexShrink: 0 }}>2</span>
                <p style={{ fontSize: '11px', color: 'rgba(255,255,255,0.9)', textTransform: 'uppercase', letterSpacing: '0.06em', margin: 0, fontWeight: 700 }}>
                  Productos e insumos
                </p>
              </div>
              <span style={{ fontSize: '12px', color: 'rgba(255,255,255,0.7)' }}>
                {grupos.length} producto(s)
              </span>
            </div>
            <div style={{ padding: '14px 16px', display: 'flex', flexDirection: 'column', gap: '12px' }}>

              {grupos.map((grupo, gIdx) => {
                const totalGrupo = grupo.filas.reduce((s, f) => s + (parseFloat(f.kg) || 0), 0)
                const completo = grupo.producto && grupo.filas.some(f => f.mp && parseFloat(f.kg) > 0)
                return (
                  <div key={grupo.id} style={{
                    borderRadius: '12px',
                    border: `2px solid ${completo ? '#004895' : '#e0e0e0'}`,
                  }}>
                    {/* Header del grupo */}
                    <div style={{
                      backgroundColor: completo ? '#004895' : '#f5f5f5',
                      padding: '8px 12px',
                      borderRadius: '10px 10px 0 0',
                      display: 'flex', alignItems: 'center', gap: '8px',
                    }}>
                      <span style={{
                        backgroundColor: completo ? 'rgba(255,255,255,0.25)' : '#ddd',
                        color: completo ? 'white' : '#555',
                        borderRadius: '50%', width: '22px', height: '22px',
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                        fontSize: '12px', fontWeight: 800, flexShrink: 0,
                      }}>P{gIdx + 1}</span>
                      <span style={{ fontSize: '13px', fontWeight: 700, color: completo ? 'white' : '#555', flex: 1 }}>
                        {grupo.producto || 'Sin producto asignado'}
                      </span>
                      {completo && (
                        <span style={{ fontSize: '12px', color: 'rgba(255,255,255,0.85)', backgroundColor: 'rgba(255,255,255,0.15)', borderRadius: '6px', padding: '2px 8px' }}>
                          {totalGrupo.toFixed(2)} kg
                        </span>
                      )}
                      {grupos.length > 1 && (
                        <button type="button" onClick={() => eliminarGrupo(grupo.id)} style={{
                          backgroundColor: completo ? 'rgba(255,255,255,0.15)' : '#ffebee',
                          color: completo ? 'rgba(255,255,255,0.9)' : '#c62828',
                          border: 'none', borderRadius: '6px', padding: '3px 8px',
                          fontSize: '11px', fontWeight: 700, cursor: 'pointer',
                        }}>✕</button>
                      )}
                    </div>

                    <div style={{ padding: '12px', display: 'flex', flexDirection: 'column', gap: '10px' }}>
                      {/* Selector de producto */}
                      <div>
                        <label style={{ ...lbl, fontSize: '12px' }}>Producto a producir *</label>
                        <SearchSelect
                          opciones={productosPA.map(p => ({
                            value: p.Nombre || p.Title || '',
                            label: (p.Nombre || p.Title || '') + (p.Codigo ? ` (${p.Codigo})` : ''),
                          }))}
                          value={grupo.producto}
                          onChange={v => updateGrupoProducto(grupo.id, v)}
                          placeholder="Buscar producto terminado..."
                        />
                      </div>

                      {/* Filas de MP */}
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        <label style={{ ...lbl, fontSize: '12px' }}>Insumos entregados *</label>
                        {grupo.filas.map((fila, fIdx) => (
                          <div key={fila.id} style={{
                            display: 'grid', gridTemplateColumns: '1fr 90px auto',
                            gap: '6px', alignItems: 'center',
                          }}>
                            <SearchSelect
                              opciones={materiasPrimas.map(mp => ({
                                value: mp.Nombre || mp.Title || '',
                                label: (mp.Nombre || mp.Title || '') + (mp.Codigo ? ` (${mp.Codigo})` : ''),
                              }))}
                              value={fila.mp}
                              onChange={v => updateFilaGrupo(grupo.id, fila.id, 'mp', v)}
                              placeholder={`Insumo ${fIdx + 1}…`}
                            />
                            <input type="number" step="0.01" min="0"
                              value={fila.kg || ''}
                              onChange={e => updateFilaGrupo(grupo.id, fila.id, 'kg', e.target.value)}
                              placeholder="kg"
                              style={{ ...inp, textAlign: 'right', fontWeight: 700, fontSize: '14px', padding: '10px 8px', boxSizing: 'border-box', borderColor: parseFloat(fila.kg) > 0 ? '#004895' : '#ddd' }} />
                            <button type="button"
                              onClick={() => eliminarFilaGrupo(grupo.id, fila.id)}
                              disabled={grupo.filas.length === 1}
                              style={{
                                width: '30px', height: '44px', border: 'none', borderRadius: '6px',
                                backgroundColor: grupo.filas.length === 1 ? 'transparent' : '#ffebee',
                                color: grupo.filas.length === 1 ? '#ddd' : '#c62828',
                                fontSize: '16px', cursor: grupo.filas.length === 1 ? 'not-allowed' : 'pointer',
                              }}>×</button>
                          </div>
                        ))}
                        <button type="button"
                          onClick={() => agregarFilaGrupo(grupo.id)}
                          disabled={grupo.filas.length >= 5}
                          style={{
                            backgroundColor: 'transparent',
                            color: grupo.filas.length >= 5 ? '#bbb' : '#004895',
                            border: `1.5px dashed ${grupo.filas.length >= 5 ? '#ddd' : '#004895'}`,
                            borderRadius: '8px', padding: '7px',
                            fontSize: '12px', fontWeight: 700,
                            cursor: grupo.filas.length >= 5 ? 'not-allowed' : 'pointer',
                          }}>
                          + Agregar insumo {grupo.filas.length >= 5 ? '(máx. 5)' : ''}
                        </button>
                      </div>

                      {totalGrupo > 0 && (
                        <div style={{ display: 'flex', justifyContent: 'flex-end', borderTop: '1px solid #e8e8e8', paddingTop: '8px' }}>
                          <span style={{ fontSize: '14px', fontWeight: 700, color: '#004895' }}>
                            Subtotal: {totalGrupo.toFixed(2)} kg
                          </span>
                        </div>
                      )}
                    </div>
                  </div>
                )
              })}

              {/* Botón agregar producto */}
              <button type="button" onClick={agregarGrupo}
                disabled={grupos.length >= 5}
                style={{
                  backgroundColor: 'transparent',
                  color: grupos.length >= 5 ? '#bbb' : '#37BEEC',
                  border: `2px dashed ${grupos.length >= 5 ? '#e0e0e0' : '#37BEEC'}`,
                  borderRadius: '12px', padding: '12px',
                  fontSize: '14px', fontWeight: 700,
                  cursor: grupos.length >= 5 ? 'not-allowed' : 'pointer',
                  display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px',
                }}>
                + Agregar otro producto {grupos.length >= 5 ? '(máx. 5)' : `(${grupos.length}/5)`}
              </button>

              {/* Total general */}
              {grupos.some(g => g.filas.some(f => parseFloat(f.kg) > 0)) && (
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', backgroundColor: '#004895', borderRadius: '10px', padding: '12px 16px' }}>
                  <span style={{ fontSize: '13px', color: 'rgba(255,255,255,0.8)' }}>
                    {grupos.reduce((s, g) => s + g.filas.filter(f => f.mp && parseFloat(f.kg) > 0).length, 0)} insumo(s) · {grupos.filter(g => g.producto).length} producto(s)
                  </span>
                  <span style={{ fontSize: '18px', fontWeight: 800, color: 'white' }}>
                    Total: {grupos.reduce((s, g) => s + g.filas.reduce((ss, f) => ss + (parseFloat(f.kg) || 0), 0), 0).toFixed(2)} kg
                  </span>
                </div>
              )}
            </div>
          </div>

          {/* Sección 3: Observaciones + Envío */}
          <div style={{ backgroundColor: 'white', borderRadius: '14px', overflow: 'hidden', boxShadow: '0 2px 8px rgba(0,0,0,0.08)' }}>
            <div style={{ backgroundColor: '#004895', padding: '10px 16px', display: 'flex', alignItems: 'center', gap: '10px' }}>
              <span style={{ backgroundColor: 'rgba(255,255,255,0.2)', borderRadius: '50%', width: '22px', height: '22px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '12px', fontWeight: 800, color: 'white', flexShrink: 0 }}>3</span>
              <p style={{ fontSize: '11px', color: 'rgba(255,255,255,0.9)', textTransform: 'uppercase', letterSpacing: '0.06em', margin: 0, fontWeight: 700 }}>Confirmar entrega</p>
            </div>
            <div style={{ padding: '14px 16px', display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div>
                <label style={lbl}>Observaciones</label>
                <textarea value={obs} onChange={e => setObs(e.target.value)} rows={2}
                  placeholder="Notas adicionales sobre la entrega..."
                  style={{ ...inp, resize: 'none', fontFamily: 'inherit' }} />
              </div>

              {feedback && (
                <div style={{
                  padding: '12px', borderRadius: '8px',
                  backgroundColor: feedback.tipo === 'exito' ? '#e3f7fd' : '#ffebee',
                  color: feedback.tipo === 'exito' ? '#0288d1' : '#c62828',
                  fontSize: '14px', fontWeight: 600, border: `1px solid ${feedback.tipo === 'exito' ? '#37BEEC' : '#f44336'}`,
                }}>
                  {feedback.tipo === 'exito' ? '✓ ' : '✗ '}{feedback.msg}
                </div>
              )}

              <button type="submit" disabled={enviando} style={{
                backgroundColor: enviando ? '#ccc' : '#004895', color: 'white',
                border: 'none', borderRadius: '12px', padding: '16px',
                fontSize: '16px', fontWeight: 700, minHeight: '56px',
                cursor: enviando ? 'not-allowed' : 'pointer',
                boxShadow: enviando ? 'none' : '0 4px 14px rgba(0,72,149,0.3)',
              }}>
                {enviando ? '⏳ Guardando...' : `✓ Registrar entregas en Kardex`}
              </button>
            </div>
          </div>

        </form>
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

                  // Transferencia pendiente (operario lo marcó al cerrar turno)
                  const tieneTransfPendiente = grupo.numeroOF && ofsConTransferencia.has(grupo.numeroOF)

                  // Colores según estado
                  const bgHeader = tieneOpSinValidar ? '#e65100' : tieneTransfPendiente ? '#0277bd' : '#004895'
                  const borderColor = tieneOpSinValidar ? '#ff9800' : tieneTransfPendiente ? '#81d4fa' : '#90caf9'

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
                        {/* Fila 1: máquina + turno + kg + flecha */}
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                            <strong style={{ fontSize: '16px' }}>{grupo.maquina}</strong>
                            <span style={{ fontSize: '12px', opacity: 0.8, backgroundColor: 'rgba(255,255,255,0.15)', borderRadius: '6px', padding: '1px 8px' }}>
                              Turno {grupo.turno}
                            </span>
                          </div>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                            <strong style={{ fontSize: '18px' }}>{neto.toFixed(2)} kg</strong>
                            <span style={{ opacity: 0.6, fontSize: '16px' }}>{abierto ? '▲' : '▼'}</span>
                          </div>
                        </div>
                        {/* Fila 1b: fecha sola */}
                        <div style={{ marginTop: '3px' }}>
                          <span style={{ fontSize: '12px', opacity: 0.7 }}>
                            {grupo.fecha ? format(new Date(grupo.fecha + 'T12:00:00'), 'dd/MM/yyyy') : '—'}
                          </span>
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
                          {tieneTransfPendiente && (
                            <span style={{
                              backgroundColor: 'rgba(255,255,255,0.2)',
                              border: '1px solid rgba(255,255,255,0.5)',
                              borderRadius: '10px', padding: '2px 10px',
                              fontSize: '11px', fontWeight: 700,
                            }}>→ Transferencia pendiente</span>
                          )}
                        </div>
                      </div>

                      {/* Items del grupo — agrupados por producto, luego por insumo */}
                      {abierto && (
                        <div style={{ backgroundColor: 'white' }}>
                          {(() => {
                            // Agrupar por producto primero
                            const porProducto = {}
                            grupo.items.forEach(e => {
                              const prodKey = (e.Producto || '').trim() || '(sin producto)'
                              if (!porProducto[prodKey]) porProducto[prodKey] = []
                              porProducto[prodKey].push(e)
                            })
                            const productosKeys = Object.keys(porProducto)
                            return productosKeys.map((prodNombre, pIdx) => {
                              const itemsProd = porProducto[prodNombre]
                              // Agrupar por insumo dentro del producto
                              const porInsumo = {}
                              itemsProd.forEach(e => {
                                const key = (e.Insumo || '—').trim().toLowerCase()
                                if (!porInsumo[key]) porInsumo[key] = []
                                porInsumo[key].push(e)
                              })
                              const totalProd = itemsProd.reduce((s, e) => s + (e.KgEntregados || 0), 0)
                              return (
                                <div key={prodNombre} style={{ borderTop: pIdx > 0 ? '2px solid #e8f0fb' : 'none' }}>
                                  {/* Sub-cabecera de producto */}
                                  <div style={{ backgroundColor: '#f0f4ff', padding: '6px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                    <span style={{ fontSize: '12px', fontWeight: 700, color: '#004895' }}>
                                      {prodNombre === '(sin producto)' ? '—' : prodNombre}
                                    </span>
                                    <span style={{ fontSize: '12px', color: '#555' }}>{totalProd.toFixed(2)} kg</span>
                                  </div>
                                  {Object.entries(porInsumo).map(([, entries], gIdx) => {
                                    const nombre = entries[0].Insumo || '—'
                                    const totalKg = entries.reduce((s, e) => {
                                      const pendiente = e.Observacion?.includes('operario') && !yaValidado(e.Observacion)
                                      const kg = pendiente ? (e.KgDeclaradoOperario || e.KgEntregados || 0) : (e.KgEntregados || 0)
                                      return s + kg
                                    }, 0)
                                    const totalDev = entries.reduce((s, e) => s + (e.KgDevueltos || 0), 0)
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
                                    <span style={{ fontSize: '16px', fontWeight: 700, color: '#004895' }}>
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
                                                style={{ backgroundColor: '#e8f0fb', color: '#004895', border: '1px solid #90caf9', borderRadius: '4px', padding: '1px 8px', fontSize: '10px', fontWeight: 700, cursor: 'pointer' }}>✓ Validar</button>
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
                                            <button onClick={() => guardarEdicionEntrada(e.ID)} disabled={guardandoEntrada} style={{ backgroundColor: '#004895', color: 'white', border: 'none', borderRadius: '6px', padding: '5px 12px', fontSize: '12px', fontWeight: 700, cursor: 'pointer' }}>{guardandoEntrada ? '...' : '✓'}</button>
                                            <button onClick={() => setEditEntradaId(null)} style={{ background: 'none', border: '1px solid #ddd', borderRadius: '6px', padding: '5px 10px', fontSize: '12px', color: '#666', cursor: 'pointer' }}>✕</button>
                                          </div>
                                        )}
                                        {/* Devolución inline */}
                                        {editDevId === e.ID && (
                                          <div style={{ display: 'flex', gap: '6px', alignItems: 'center', marginTop: '6px' }}>
                                            <span style={{ fontSize: '11px', color: '#555', whiteSpace: 'nowrap' }}>Kg devueltos:</span>
                                            <input type="number" step="0.01" min="0" value={editDevKg} onChange={ev => setEditDevKg(ev.target.value)} autoFocus style={{ ...inp, flex: 1, padding: '5px 8px', textAlign: 'right', fontWeight: 700, fontSize: '13px' }} />
                                            <button onClick={() => guardarDevolucion(e.ID)} disabled={guardandoDev} style={{ backgroundColor: '#004895', color: 'white', border: 'none', borderRadius: '6px', padding: '5px 12px', fontSize: '12px', fontWeight: 700, cursor: 'pointer' }}>{guardandoDev ? '...' : '✓'}</button>
                                            <button onClick={() => setEditDevId(null)} style={{ background: 'none', border: '1px solid #ddd', borderRadius: '6px', padding: '5px 10px', fontSize: '12px', color: '#666', cursor: 'pointer' }}>✕</button>
                                          </div>
                                        )}
                                      </div>
                                    )
                                  })}
                                </div>
                              )
                            })}
                          </div>
                        )
                      })
                    })()}
                          {/* Acciones del grupo */}
                          <div style={{ padding: '10px 14px', borderTop: '1px solid #f0f0f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
                            <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
                              {(() => {
                                const yaProdujo = grupo.numeroOF && ofsConRegistro.has(grupo.numeroOF)
                                return (
                                  <button
                                    onClick={() => { if (!yaProdujo) { setModalReasignar({ grupo }); setMaqDestinoSel('') } }}
                                    disabled={yaProdujo}
                                    title={yaProdujo ? 'Ya hay producción registrada — usa "Transferir kg restantes"' : 'Corregir máquina asignada'}
                                    style={{
                                      backgroundColor: yaProdujo ? '#f5f5f5' : '#f0f4ff',
                                      color: yaProdujo ? '#bbb' : '#004895',
                                      border: `1.5px solid ${yaProdujo ? '#e0e0e0' : '#004895'}`,
                                      borderRadius: '8px', padding: '5px 12px', fontSize: '11px', fontWeight: 700,
                                      cursor: yaProdujo ? 'not-allowed' : 'pointer',
                                    }}>
                                    ↔ Reasignar máquina
                                  </button>
                                )
                              })()}
                              {(() => {
                                const hayMpUsada = grupo.items.some(e => (e.KgUsado || 0) > 0)
                                return (
                                  <button
                                    onClick={() => { if (hayMpUsada) { setModalTransferir({ grupo }); setMaqDestinoSel('') } }}
                                    disabled={!hayMpUsada}
                                    title={!hayMpUsada ? 'Aún no se usó MP — usa "Reasignar máquina"' : 'Transferir kg no usados a otra máquina'}
                                    style={{
                                      backgroundColor: !hayMpUsada ? '#f5f5f5' : tieneTransfPendiente ? '#37BEEC' : '#e3f7fd',
                                      color: !hayMpUsada ? '#bbb' : tieneTransfPendiente ? 'white' : '#0288d1',
                                      border: `1.5px solid ${!hayMpUsada ? '#e0e0e0' : tieneTransfPendiente ? '#37BEEC' : '#0288d1'}`,
                                      borderRadius: '8px', padding: '5px 12px', fontSize: '11px', fontWeight: 700,
                                      cursor: !hayMpUsada ? 'not-allowed' : 'pointer',
                                    }}>
                                    → Transferir kg restantes
                                  </button>
                                )
                              })()}
                            </div>
                            <button
                              onClick={() => {
                                setModalAdicional({ maquina: grupo.maquina, turno: grupo.turno, fecha: grupo.fecha, numeroOF: grupo.numeroOF })
                                setFilasAdicionales([filaVacia()])
                              }}
                              style={{ backgroundColor: 'transparent', color: '#004895', border: '1.5px solid #004895', borderRadius: '8px', padding: '5px 14px', fontSize: '11px', fontWeight: 700, cursor: 'pointer' }}>
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

      {/* Modal Reasignar máquina */}
      {modalReasignar && (
        <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 300, padding: '16px' }}>
          <div style={{ backgroundColor: 'white', borderRadius: '16px', padding: '24px', width: '100%', maxWidth: '420px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
              <div>
                <h3 style={{ color: '#004895', fontSize: '16px', fontWeight: 700, margin: 0 }}>↔ Reasignar máquina</h3>
                <p style={{ fontSize: '12px', color: '#888', marginTop: '4px' }}>
                  Todos los registros de esta OF se moverán a la máquina seleccionada.
                </p>
              </div>
              <button onClick={() => setModalReasignar(null)} style={{ background: 'none', border: 'none', fontSize: '22px', cursor: 'pointer', color: '#666' }}>✕</button>
            </div>

            <div style={{ backgroundColor: '#f0f4ff', borderRadius: '10px', padding: '12px' }}>
              <p style={{ fontSize: '12px', color: '#555', margin: '0 0 4px' }}>Grupo actual</p>
              <p style={{ fontSize: '14px', fontWeight: 700, color: '#1a1a1a', margin: 0 }}>
                {modalReasignar.grupo.maquina} · Turno {modalReasignar.grupo.turno}
              </p>
              {modalReasignar.grupo.numeroOF && (
                <p style={{ fontSize: '11px', color: '#888', margin: '4px 0 0', fontFamily: 'monospace' }}>
                  {modalReasignar.grupo.numeroOF}
                </p>
              )}
              <p style={{ fontSize: '12px', color: '#555', margin: '8px 0 0' }}>
                {modalReasignar.grupo.items.length} entrada(s) · {modalReasignar.grupo.items.reduce((s, i) => s + (i.KgEntregados || 0), 0).toFixed(2)} kg entregados
              </p>
            </div>

            <div>
              <label style={lbl}>Máquina destino *</label>
              <select value={maqDestinoSel} onChange={e => setMaqDestinoSel(e.target.value)} style={sel}>
                <option value="">Seleccionar máquina...</option>
                {maquinas.filter(m => (m.Codigo || m.Nombre || m.Title) !== modalReasignar.grupo.maquina).map(m => {
                  const nombre = m.Nombre || m.Title || ''
                  const codigo = m.Codigo || ''
                  return <option key={m.ID} value={codigo || nombre}>{nombre}{codigo ? ` (${codigo})` : ''}</option>
                })}
              </select>
            </div>

            <div style={{ backgroundColor: '#fff8e1', borderRadius: '8px', padding: '10px 12px', fontSize: '12px', color: '#e65100' }}>
              ⚠ Esta acción mueve <strong>todos</strong> los registros Kardex de la OF a la máquina destino. Usa esta opción solo cuando la asignación original fue incorrecta.
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', gap: '10px' }}>
              <button onClick={() => setModalReasignar(null)} style={{ padding: '12px', borderRadius: '10px', border: '2px solid #ddd', background: 'white', color: '#555', fontSize: '14px', cursor: 'pointer' }}>Cancelar</button>
              <button onClick={ejecutarReasignar} disabled={!maqDestinoSel || guardandoReasig} style={{
                padding: '12px', borderRadius: '10px', border: 'none',
                backgroundColor: !maqDestinoSel || guardandoReasig ? '#ccc' : '#004895',
                color: 'white', fontSize: '14px', fontWeight: 700,
                cursor: !maqDestinoSel || guardandoReasig ? 'not-allowed' : 'pointer',
              }}>
                {guardandoReasig ? '⏳ Reasignando...' : '↔ Confirmar reasignación'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Transferir kg restantes */}
      {modalTransferir && (() => {
        const grupo = modalTransferir.grupo
        const porInsumo = {}
        grupo.items.forEach(e => {
          const key = (e.Insumo || '').trim().toLowerCase()
          if (!porInsumo[key]) porInsumo[key] = { nombre: e.Insumo, disponible: 0 }
          porInsumo[key].disponible += Math.max(0,
            (e.KgEntregados || 0) - (e.KgUsado || 0) - (e.KgMermaRec || 0)
            - (e.KgMermaNoRec || 0) - (e.KgDevueltos || 0)
          )
        })
        const insumosDisp = Object.values(porInsumo).filter(i => i.disponible > 0.001)
        const totalKgDisp = insumosDisp.reduce((s, i) => s + i.disponible, 0)
        return (
          <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 300, padding: '16px' }}>
            <div style={{ backgroundColor: 'white', borderRadius: '16px', padding: '24px', width: '100%', maxWidth: '420px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                <div>
                  <h3 style={{ color: '#0288d1', fontSize: '16px', fontWeight: 700, margin: 0 }}>→ Transferir kg restantes</h3>
                  <p style={{ fontSize: '12px', color: '#888', marginTop: '4px' }}>
                    Los kg no usados se asignan a otra máquina con la misma OF.
                  </p>
                </div>
                <button onClick={() => setModalTransferir(null)} style={{ background: 'none', border: 'none', fontSize: '22px', cursor: 'pointer', color: '#666' }}>✕</button>
              </div>

              <div style={{ backgroundColor: '#e3f7fd', borderRadius: '10px', padding: '12px' }}>
                <p style={{ fontSize: '12px', color: '#555', margin: '0 0 6px' }}>Kg disponibles para transferir</p>
                {insumosDisp.length === 0 ? (
                  <p style={{ fontSize: '13px', color: '#e65100', margin: 0 }}>⚠ No hay kg disponibles — todos fueron usados o devueltos.</p>
                ) : insumosDisp.map(i => (
                  <div key={i.nombre} style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '4px' }}>
                    <span style={{ fontSize: '13px', color: '#1a1a1a' }}>• {i.nombre}</span>
                    <strong style={{ fontSize: '13px', color: '#0288d1' }}>{i.disponible.toFixed(2)} kg</strong>
                  </div>
                ))}
                {insumosDisp.length > 0 && (
                  <div style={{ borderTop: '1px solid #b3e5fc', marginTop: '6px', paddingTop: '6px', display: 'flex', justifyContent: 'space-between' }}>
                    <strong style={{ fontSize: '13px', color: '#555' }}>Total</strong>
                    <strong style={{ fontSize: '14px', color: '#0288d1' }}>{totalKgDisp.toFixed(2)} kg</strong>
                  </div>
                )}
              </div>

              <div>
                <label style={lbl}>Máquina destino *</label>
                <select value={maqDestinoSel} onChange={e => setMaqDestinoSel(e.target.value)} style={sel}>
                  <option value="">Seleccionar máquina...</option>
                  {maquinas.filter(m => (m.Codigo || m.Nombre || m.Title) !== grupo.maquina).map(m => {
                    const nombre = m.Nombre || m.Title || ''
                    const codigo = m.Codigo || ''
                    return <option key={m.ID} value={codigo || nombre}>{nombre}{codigo ? ` (${codigo})` : ''}</option>
                  })}
                </select>
              </div>

              {maqDestinoSel && insumosDisp.length > 0 && (
                <div style={{ backgroundColor: '#f0f4ff', borderRadius: '8px', padding: '10px 12px', fontSize: '12px', color: '#555' }}>
                  Se crearán <strong>{insumosDisp.length} entrada(s)</strong> en <strong>{maqDestinoSel}</strong> con la OF <span style={{ fontFamily: 'monospace' }}>{grupo.numeroOF}</span>.
                  Los kg en <strong>{grupo.maquina}</strong> quedarán registrados como transferidos.
                </div>
              )}

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', gap: '10px' }}>
                <button onClick={() => setModalTransferir(null)} style={{ padding: '12px', borderRadius: '10px', border: '2px solid #ddd', background: 'white', color: '#555', fontSize: '14px', cursor: 'pointer' }}>Cancelar</button>
                <button onClick={ejecutarTransferir} disabled={!maqDestinoSel || insumosDisp.length === 0 || guardandoTransf} style={{
                  padding: '12px', borderRadius: '10px', border: 'none',
                  backgroundColor: !maqDestinoSel || insumosDisp.length === 0 || guardandoTransf ? '#ccc' : '#37BEEC',
                  color: 'white', fontSize: '14px', fontWeight: 700,
                  cursor: !maqDestinoSel || insumosDisp.length === 0 || guardandoTransf ? 'not-allowed' : 'pointer',
                }}>
                  {guardandoTransf ? '⏳ Transfiriendo...' : `→ Transferir ${totalKgDisp.toFixed(2)} kg`}
                </button>
              </div>
            </div>
          </div>
        )
      })()}

      {/* Modal MP adicional a grupo existente */}
      {modalAdicional && (
        <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 300, padding: '16px' }}>
          <div style={{ backgroundColor: 'white', borderRadius: '16px', padding: '24px', width: '100%', maxWidth: '460px', display: 'flex', flexDirection: 'column', gap: '16px', maxHeight: '90vh', overflowY: 'auto' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <h3 style={{ color: '#004895', fontSize: '16px', fontWeight: 700 }}>Agregar MP adicional</h3>
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
              backgroundColor: '#e8f0fb', color: '#004895', border: '1px solid #004895',
              borderRadius: '8px', padding: '8px', fontSize: '13px', fontWeight: 700, cursor: 'pointer',
            }}>+ Agregar otra MP</button>

            {filasAdicionales.some(f => parseFloat(f.kg) > 0) && (
              <p style={{ fontSize: '13px', fontWeight: 700, color: '#004895', textAlign: 'right' }}>
                Total: {filasAdicionales.reduce((s, f) => s + (parseFloat(f.kg) || 0), 0).toFixed(2)} kg
              </p>
            )}

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', gap: '10px' }}>
              <button onClick={() => setModalAdicional(null)} style={{ padding: '12px', borderRadius: '10px', border: '2px solid #ccc', backgroundColor: 'white', color: '#555', fontSize: '14px', cursor: 'pointer' }}>Cancelar</button>
              <button onClick={guardarAdicionales} disabled={guardandoAd} style={{
                padding: '12px', borderRadius: '10px', border: 'none',
                backgroundColor: guardandoAd ? '#ccc' : '#004895', color: 'white',
                fontSize: '14px', fontWeight: 700, cursor: guardandoAd ? 'not-allowed' : 'pointer',
              }}>{guardandoAd ? '⏳ Guardando...' : '✓ Guardar MP adicional'}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
