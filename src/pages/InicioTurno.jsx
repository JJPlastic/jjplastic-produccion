import { useState, useEffect } from 'react'
import { useForm } from 'react-hook-form'
import { format } from 'date-fns'
import { es } from 'date-fns/locale'
import { useMsal } from '../hooks/useMsal'
import { useMaestros } from '../hooks/useMaestros'
import { useApp } from '../context/AppContext'
import { Header } from '../components/Header'
import { LoadingSpinner, FullscreenFeedback } from '../components/LoadingSpinner'
import { SearchSelect } from '../components/SearchSelect'
import { createListItem, updateListItem, getListItems, getOFsActivas } from '../services/sharepoint'
import { encolarOperacion } from '../services/indexedDB'
import { tabletConfig } from '../config/tabletConfig'

const TURNOS = [
  { id: 'M', label: 'Mañana', hora: '6:00 – 14:00' },
  { id: 'T', label: 'Tarde', hora: '14:00 – 22:00' },
  { id: 'N', label: 'Noche', hora: '22:00 – 6:00' },
]

const Field = ({ label, error, children }) => (
  <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
    <label style={{ fontWeight: 600, fontSize: '14px', color: '#333' }}>{label}</label>
    {children}
    {error && <p style={{ color: '#d32f2f', fontSize: '13px' }}>{error}</p>}
  </div>
)

const selectStyle = {
  width: '100%', padding: '14px 12px', borderRadius: '10px',
  border: '2px solid #ddd', fontSize: '16px',
  backgroundColor: 'white', color: '#1a1a1a',
  minHeight: '52px', appearance: 'auto',
}

const inputStyle = {
  ...selectStyle,
  color: '#1a1a1a',
}

export default function InicioTurno() {
  const { getToken, usuario, logout } = useMsal()
  const { setTurnoActivo, pendingCount, seleccionarRol, setPantalla, ofPreseleccionada, setOfPreseleccionada } = useApp()
  const { operarios, productos, colores, cargando } = useMaestros(getToken)

  // Cargar MP directamente desde SP (independiente del cache de maestros)
  const [materiasPrimas, setMateriasPrimas] = useState([])
  useEffect(() => {
    let cancelled = false
    const cargarMP = async () => {
      try {
        const token = await getToken()
        if (!token) return
        const todos = await getListItems(token, 'Maestro_Productos', { top: 500 })
        if (cancelled) return
        const esActivo = p => p.Activo === true || p.Activo === 1 || p.Activo === undefined
        const mp = todos.filter(p => esActivo(p) && ['MP','MC'].includes((p.TipoProducto || '').trim().toUpperCase()))
        setMateriasPrimas(mp)
      } catch (err) {
        console.error('Error cargando MP:', err)
      }
    }
    if (navigator.onLine) cargarMP()
    return () => { cancelled = true }
  }, [])
  const { register, handleSubmit, setValue, watch, formState: { errors } } = useForm()
  const [enviando, setEnviando]             = useState(false)
  const [feedback, setFeedback]             = useState(null)
  const [ofsActivas, setOfsActivas]         = useState([])
  const [ofSeleccionada, setOfSeleccionada] = useState(null)
  // Filas de MP (múltiples materias primas)
  const [filasMP, setFilasMP] = useState([{ id: crypto.randomUUID(), mp: '', kg: '' }])
  const agregarFilaMP  = () => setFilasMP(p => [...p, { id: crypto.randomUUID(), mp: '', kg: '' }])
  const eliminarFilaMP = (id) => setFilasMP(p => p.length > 1 ? p.filter(f => f.id !== id) : p)
  const updateFilaMP   = (id, campo, val) => setFilasMP(p => p.map(f => f.id === id ? { ...f, [campo]: val } : f))

  const turnoSel = watch('Turno')

  // Auto-detectar turno según la hora actual
  useEffect(() => {
    const hora = new Date().getHours()
    const turnoAuto = hora >= 6 && hora < 14 ? 'M'
                    : hora >= 14 && hora < 22 ? 'T'
                    : 'N'
    setValue('Turno', turnoAuto, { shouldValidate: false })
  }, [])

  // Cargar OFs activas para esta máquina al montar
  useEffect(() => {
    let cancelled = false
    const cargarOFs = async () => {
      try {
        const token = await getToken()
        if (!token) return
        const ofs = await getOFsActivas(token, tabletConfig.codigoMaquina)
        if (!cancelled) setOfsActivas(ofs)
      } catch { /* sin OF activas */ }
    }
    if (navigator.onLine) cargarOFs()
    return () => { cancelled = true }
  }, [])

  // Si viene con OF preseleccionada desde Bienvenida → auto-seleccionarla
  useEffect(() => {
    if (!ofPreseleccionada) return
    // Usar el objeto directo (ya tiene mp[], tieneKardex, etc.)
    setOfSeleccionada(ofPreseleccionada)
    setOfPreseleccionada(null) // consumir para no repetir
  }, [ofPreseleccionada])

  // Cuando operario selecciona una OF → pre-llenar filas de MP desde Kardex
  useEffect(() => {
    if (!ofSeleccionada) return
    // Usar los datos de MP ya cargados en la OF (vienen de getOFsActivas)
    if (ofSeleccionada.mp && ofSeleccionada.mp.length > 0) {
      // Consolidar entradas del mismo insumo (suma kg) antes de crear filas
      const consolidado = {}
      ofSeleccionada.mp
        .filter(m => m.insumo && m.insumo !== '—' && m.insumo.trim() !== '')
        .forEach(m => {
          const key = m.insumo.trim().toLowerCase()
          if (!consolidado[key]) consolidado[key] = { insumo: m.insumo, kg: 0 }
          consolidado[key].kg += m.kg || 0
        })
      const filasKardex = Object.values(consolidado).map(m => ({
        id: crypto.randomUUID(),
        mp: m.insumo,
        kg: String(m.kg > 0 ? m.kg : ''),
      }))
      setFilasMP(filasKardex.length > 0 ? filasKardex : [{ id: crypto.randomUUID(), mp: '', kg: '' }])
    } else {
      // OF sin Kardex (operario registró primero) → dejar fila vacía
      setFilasMP([{ id: crypto.randomUUID(), mp: '', kg: '' }])
    }
  }, [ofSeleccionada?.of])

  // Auto-seleccionar operario cuyo Email coincida con la cuenta M365 logueada
  useEffect(() => {
    if (!operarios.length || !usuario?.email) return
    const match = operarios.find(o =>
      (o.Email || '').toLowerCase() === usuario.email.toLowerCase()
    )
    if (match) setValue('Operario', String(match.ID), { shouldValidate: false })
  }, [operarios, usuario?.email])

  const onSubmit = async (data) => {
    setEnviando(true)
    const ahora = new Date()

    // Buscar objetos completos por ID del maestro local
    const opObj = operarios.find(o => String(o.ID) === String(data.Operario))
    const apObj = data.OperarioApoyo ? operarios.find(o => String(o.ID) === String(data.OperarioApoyo)) : null
    const opNombre = opObj ? (opObj.Nombre || opObj.Title || '') : ''
    const apNombre = apObj ? (apObj.Nombre || apObj.Title || '') : ''

    // KgPorUnidad del producto seleccionado — necesario para calcular semáforo al cerrar
    const prodObj = productos.find(p => (p.Codigo || '') === data.Producto)
    const kgPorUnidad    = prodObj?.KgPorUnidad ?? 0
    const productoNombre = prodObj?.Nombre || prodObj?.Title || data.Producto

    // Código de lote: por producto
    const codigoLote = `${tabletConfig.codigoMaquina}-${data.Producto}-${format(ahora, 'yyyyMMdd-HHmm')}`
    // Número OF: agrupa todos los productos de la misma sesión de máquina
    // Si operario seleccionó una OF activa (PCP registró Kardex primero) → usarla
    // Si no → generar nueva (sin código de producto, cubre toda la sesión)
    const numeroOF = ofSeleccionada
      ? ofSeleccionada.of
      : `OF-${tabletConfig.codigoMaquina}-${format(ahora, 'yyyyMMdd-HHmm')}`

    // Objeto local para display (guarda nombres legibles + datos para cálculos de cierre)
    const registroLocal = {
      Maquina: tabletConfig.codigoMaquina,
      Turno: data.Turno,
      Operario: opNombre,
      Operario_Apoyo: apNombre || null,
      Producto: data.Producto,
      Color: data.Color,
      HoraInicio: ahora.toISOString(),
      Fecha: ahora.toISOString(),
      Estado: 'abierto',
      CheckboxVeracidad: false,
      Tipo_Apertura: 'Normal',
      Registro_Previo_ID: null,
      Paradas: '[]',
      Estado_Validacion: 'Verde',
      KgPorUnidadProducto: kgPorUnidad,
      ProductoNombre: productoNombre,
      Codigo_Lote: codigoLote,
      Numero_OF: numeroOF,
    }

    // Guardar cada fila de MP en Kardex_MP
    const filasValidasMP = filasMP.filter(f => f.mp && parseFloat(f.kg) > 0)

    try {
      const token = await getToken()
      if (token && navigator.onLine) {
        // ── Validación anti-duplicado: verificar que no exista registro abierto
        // para esta máquina en el mismo turno y fecha
        const hoy = ahora.toISOString().split('T')[0]
        const existentes = await getListItems(token, 'Registro_Produccion', {
          filter: `Title eq '${tabletConfig.codigoMaquina}' and Turno eq '${registroLocal.Turno}' and Estado eq 'abierto'`,
        })
        const duplicado = existentes.some(r => (r.Fecha || r.HoraInicio || '').startsWith(hoy))
        if (duplicado) {
          setEnviando(false)
          alert(`⚠ Ya existe un registro abierto para ${tabletConfig.nombreMaquina} en el turno ${registroLocal.Turno} de hoy. Ciérralo antes de abrir uno nuevo.`)
          return
        }

        // Payload SP: Operario y Operario_Apoyo son ahora texto simple (no Person)
        const registroSP = {
          Title: tabletConfig.codigoMaquina,
          Turno: registroLocal.Turno,
          Operario: opNombre,
          Operario_Apoyo: apNombre || null,
          Producto: registroLocal.Producto,
          Color: registroLocal.Color,
          HoraInicio: registroLocal.HoraInicio,
          Fecha: registroLocal.Fecha,
          Estado: registroLocal.Estado,
          CheckboxVeracidad: registroLocal.CheckboxVeracidad,
          Tipo_Apertura: registroLocal.Tipo_Apertura,
          Paradas: registroLocal.Paradas,
          Estado_Validacion: registroLocal.Estado_Validacion,
          Codigo_Lote: registroLocal.Codigo_Lote,
          Numero_OF: registroLocal.Numero_OF,
          ...(registroLocal.Registro_Previo_ID ? { Registro_Previo_ID: registroLocal.Registro_Previo_ID } : {}),
        }

        const resultado = await createListItem(token, 'Registro_Produccion', registroSP)
        const nuevoSpId = resultado.ID

        if (filasValidasMP.length > 0) {
          if (!ofSeleccionada?.tieneKardex) {
            // Operario va primero → crear entradas nuevas en Kardex
            await Promise.all(filasValidasMP.map(f =>
              createListItem(token, 'Kardex_MP', {
                Title: tabletConfig.codigoMaquina,
                Fecha: registroLocal.Fecha,
                Turno: registroLocal.Turno,
                Insumo: f.mp,
                KgEntregados: parseFloat(f.kg),
                KgDeclaradoOperario: parseFloat(f.kg),
                KgDevueltos: 0,
                Observacion: 'Registrada por operario',
                Numero_OF: numeroOF,
              })
            ))
          } else {
            // PCP fue primero → guardar lo que el operario declaró recibir
            // Cargamos las entradas actuales para hacer el match por Insumo
            const kardexActual = await getListItems(token, 'Kardex_MP', { top: 200 })
            const kardexOF = kardexActual.filter(k => k.Numero_OF === numeroOF)

            await Promise.all(filasValidasMP.map(async f => {
              const matching = kardexOF.find(k =>
                (k.Insumo || '').toLowerCase() === (f.mp || '').toLowerCase()
              )
              if (matching) {
                const kgOp = parseFloat(f.kg)
                const hayDiferencia = Math.abs(kgOp - (matching.KgEntregados || 0)) > 0.01
                await updateListItem(token, 'Kardex_MP', matching.ID, {
                  KgDeclaradoOperario: kgOp,
                  // Si el operario declara diferente a lo que PCP registró → marcar para validación
                  ...(hayDiferencia ? { Observacion: 'Registrada por operario' } : {}),
                })
              } else {
                // MP que el operario recibió pero PCP no registró → crear nueva entrada
                await createListItem(token, 'Kardex_MP', {
                  Title: tabletConfig.codigoMaquina,
                  Fecha: registroLocal.Fecha,
                  Turno: registroLocal.Turno,
                  Insumo: f.mp,
                  KgEntregados: 0,
                  KgDeclaradoOperario: parseFloat(f.kg),
                  KgDevueltos: 0,
                  Observacion: 'Declarada por operario — no en Kardex PCP',
                  Numero_OF: numeroOF,
                })
              }
            }))
          }
        }

        await setTurnoActivo({ ...registroLocal, spId: nuevoSpId, localId: crypto.randomUUID() })
      } else {
        const localId = crypto.randomUUID()
        await encolarOperacion({ tipo: 'create', listName: 'Registro_Produccion', data: registroLocal, localId })
        await setTurnoActivo({ ...registroLocal, spId: null, localId })
      }
      setFeedback('exito')
    } catch (err) {
      console.error('Error al iniciar turno:', err)
      const localId = crypto.randomUUID()
      await encolarOperacion({ tipo: 'create', listName: 'Registro_Produccion', data: registroLocal, localId })
      await setTurnoActivo({ ...registroLocal, spId: null, localId })
      setFeedback('exito')
    } finally {
      setEnviando(false)
    }
  }

  if (cargando) return <LoadingSpinner mensaje="Cargando datos..." />

  if (feedback === 'exito') {
    return (
      <FullscreenFeedback
        tipo="exito"
        mensaje="¡Turno iniciado!"
        onContinuar={() => setFeedback(null)}
      />
    )
  }

  return (
    <div style={{ backgroundColor: '#f0f2f5', minHeight: '100vh' }}>
      <Header
        titulo="Inicio de turno"
        subtitulo={format(new Date(), "EEEE d 'de' MMMM", { locale: es })}
        pendingCount={pendingCount}
        onLogout={logout}
        onCambiarRol={() => seleccionarRol(null)}
      />

      <div style={{ padding: '10px 12px', maxWidth: '480px', margin: '0 auto' }}>
        <form onSubmit={handleSubmit(onSubmit)} style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>

          {/* OF — si viene preseleccionada */}
          {ofSeleccionada && (
            <div style={{ backgroundColor: '#e8f5e9', borderRadius: '8px', padding: '7px 12px' }}>
              <span style={{ color: '#1b5e20', fontWeight: 700, fontSize: '12px', fontFamily: 'monospace' }}>✓ {ofSeleccionada.of}</span>
            </div>
          )}

          {/* Turno + Operario + Apoyo en una sola tarjeta */}
          <div style={{ backgroundColor: 'white', borderRadius: '10px', padding: '10px 12px', border: '1.5px solid #e0e0e0', display: 'flex', flexDirection: 'column', gap: '8px' }}>
            {/* Turno */}
            <div>
              <label style={{ fontSize: '11px', fontWeight: 600, color: '#888', textTransform: 'uppercase', display: 'block', marginBottom: '5px' }}>Turno *</label>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '6px' }}>
                {TURNOS.map(t => (
                  <button key={t.id} type="button"
                    onClick={() => setValue('Turno', t.id, { shouldValidate: true })}
                    style={{
                      padding: '10px 4px', borderRadius: '8px', border: '2px solid',
                      borderColor: turnoSel === t.id ? '#004895' : '#ddd',
                      backgroundColor: turnoSel === t.id ? '#004895' : '#f8f8f8',
                      color: turnoSel === t.id ? 'white' : '#333',
                      fontWeight: 700, fontSize: '14px', cursor: 'pointer', lineHeight: 1.2,
                    }}>
                    <div>{t.label}</div>
                  </button>
                ))}
              </div>
              <input type="hidden" {...register('Turno', { required: 'Selecciona un turno' })} />
              {errors.Turno && <p style={{ color: '#d32f2f', fontSize: '11px', marginTop: '2px' }}>{errors.Turno.message}</p>}
            </div>
            {/* Operario + Apoyo */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
              <div>
                <label style={{ fontSize: '11px', fontWeight: 600, color: '#888', textTransform: 'uppercase', display: 'block', marginBottom: '4px' }}>Operario *</label>
                <input type="hidden" {...register('Operario', { required: 'Selecciona un operario' })} />
                <SearchSelect opciones={operarios.map(op => ({ value: String(op.ID), label: op.Nombre || op.Title || '' }))}
                  value={watch('Operario')} onChange={v => setValue('Operario', v, { shouldValidate: true })} placeholder="Buscar..." />
                {errors.Operario && <p style={{ color: '#d32f2f', fontSize: '11px', marginTop: '2px' }}>{errors.Operario.message}</p>}
              </div>
              <div>
                <label style={{ fontSize: '11px', fontWeight: 600, color: '#888', textTransform: 'uppercase', display: 'block', marginBottom: '4px' }}>Apoyo</label>
                <SearchSelect opciones={[{ value: '', label: 'Sin apoyo' }, ...operarios.map(op => ({ value: String(op.ID), label: op.Nombre || op.Title || '' }))]}
                  value={watch('OperarioApoyo') || ''} onChange={v => setValue('OperarioApoyo', v)} placeholder="Buscar..." />
                <input type="hidden" {...register('OperarioApoyo')} />
              </div>
            </div>
          </div>

          {/* Producto — línea completa */}
          <div>
            <label style={{ fontSize: '12px', fontWeight: 600, color: '#555', display: 'block', marginBottom: '4px' }}>Producto *</label>
            <input type="hidden" {...register('Producto', { required: 'Selecciona un producto' })} />
            <SearchSelect opciones={productos.map(p => {
              const nombre = p.Nombre || p.Title || ''; const codigo = p.Codigo || ''
              return { value: codigo || nombre, label: nombre + (codigo ? ` (${codigo})` : '') }
            })} value={watch('Producto')} onChange={v => setValue('Producto', v, { shouldValidate: true })} placeholder="Buscar producto..." />
            {errors.Producto && <p style={{ color: '#d32f2f', fontSize: '11px', marginTop: '2px' }}>{errors.Producto.message}</p>}
          </div>

          {/* Color — línea completa */}
          <div>
            <label style={{ fontSize: '12px', fontWeight: 600, color: '#555', display: 'block', marginBottom: '4px' }}>Color *</label>
            <input type="hidden" {...register('Color', { required: 'Selecciona un color' })} />
            <SearchSelect opciones={colores.map(c => ({ value: c.Nombre || c.Title || '', label: c.Nombre || c.Title || '' }))}
              value={watch('Color')} onChange={v => setValue('Color', v, { shouldValidate: true })} placeholder="Buscar color..." />
            {errors.Color && <p style={{ color: '#d32f2f', fontSize: '11px', marginTop: '2px' }}>{errors.Color.message}</p>}
          </div>

          {/* MP recibida */}
          <div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '5px' }}>
              <label style={{ fontWeight: 600, fontSize: '12px', color: '#555' }}>
                MP recibida *
                {ofSeleccionada?.tieneKardex && <span style={{ fontSize: '11px', color: '#2e7d32', marginLeft: '6px', fontWeight: 600 }}>✓ Kardex</span>}
              </label>
              <button type="button" onClick={agregarFilaMP} style={{
                backgroundColor: '#004895', color: 'white', border: 'none',
                borderRadius: '6px', padding: '4px 10px', fontSize: '11px', fontWeight: 700, cursor: 'pointer',
              }}>+ Agregar MP</button>
            </div>
            <div>
              {filasMP.map(fila => (
                <div key={fila.id} style={{ display: 'grid', gridTemplateColumns: '1fr 90px 28px', gap: '5px', marginBottom: '6px', alignItems: 'center' }}>
                  <SearchSelect opciones={materiasPrimas.map(mp => ({
                    value: mp.Nombre || mp.Title || '',
                    label: (mp.Nombre || mp.Title || '') + (mp.Codigo ? ` (${mp.Codigo})` : ''),
                  }))} value={fila.mp} onChange={v => updateFilaMP(fila.id, 'mp', v)} placeholder="Buscar MP..." />
                  <input type="number" step="0.01" min="0" value={fila.kg}
                    onChange={e => updateFilaMP(fila.id, 'kg', e.target.value)} placeholder="0.00"
                    style={{ ...selectStyle, textAlign: 'right', fontWeight: 700, fontSize: '15px', padding: '10px 8px' }} />
                  <button type="button" onClick={() => eliminarFilaMP(fila.id)} disabled={filasMP.length === 1}
                    style={{ backgroundColor: filasMP.length === 1 ? '#f5f5f5' : '#ffebee', color: filasMP.length === 1 ? '#ccc' : '#c62828',
                      border: 'none', borderRadius: '5px', width: '28px', height: '44px', fontSize: '15px', cursor: filasMP.length === 1 ? 'not-allowed' : 'pointer' }}>×</button>
                </div>
              ))}
            </div>
            {filasMP.some(f => parseFloat(f.kg) > 0) && (
              <div style={{ textAlign: 'right', fontSize: '13px', fontWeight: 800, color: '#004895', marginTop: '2px' }}>
                Total: {filasMP.reduce((s, f) => s + (parseFloat(f.kg) || 0), 0).toFixed(2)} kg
              </div>
            )}
          </div>

          <button type="submit" disabled={enviando} style={{
            backgroundColor: enviando ? '#ccc' : '#004895', color: 'white',
            border: 'none', borderRadius: '12px', padding: '15px',
            fontSize: '16px', fontWeight: 700, minHeight: '52px', width: '100%',
            cursor: enviando ? 'not-allowed' : 'pointer',
          }}>
            {enviando ? '⏳ Iniciando...' : 'Iniciar turno →'}
          </button>

          <button
            type="button"
            onClick={() => setPantalla('bienvenida')}
            style={{
              backgroundColor: 'transparent', color: '#555',
              border: '2px solid #ccc', borderRadius: '12px',
              padding: '14px', fontSize: '15px', cursor: 'pointer',
              width: '100%',
            }}
          >
            ← Volver
          </button>

        </form>
      </div>
    </div>
  )
}
