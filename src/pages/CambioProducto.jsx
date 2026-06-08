import { useState, useEffect } from 'react'
import { useForm } from 'react-hook-form'
import { format } from 'date-fns'
import { useMsal } from '../hooks/useMsal'
import { useApp } from '../context/AppContext'
import { useMaestros } from '../hooks/useMaestros'
import { Header } from '../components/Header'
import { SearchSelect } from '../components/SearchSelect'
import { createListItem } from '../services/sharepoint'
import { encolarOperacion } from '../services/indexedDB'
import { tabletConfig } from '../config/tabletConfig'

const inputStyle = {
  width: '100%', padding: '14px 12px', borderRadius: '10px',
  border: '2px solid #ddd', fontSize: '16px',
  backgroundColor: 'white', color: '#1a1a1a', minHeight: '52px',
}

const Field = ({ label, error, children }) => (
  <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
    <label style={{ fontWeight: 600, fontSize: '14px', color: '#333' }}>{label}</label>
    {children}
    {error && <p style={{ color: '#d32f2f', fontSize: '13px' }}>{error}</p>}
  </div>
)

export default function CambioProducto() {
  const { getToken, usuario, logout } = useMsal()
  const { turnoActivo, setTurnoActivo, pendingCount, limpiarTurno, modoRelevo, setModoRelevo, setPantalla } = useApp()
  const { productos, colores, operarios, cargando } = useMaestros(getToken)

  const { register, handleSubmit, setValue, formState: { errors } } = useForm()
  const [enviando, setEnviando] = useState(false)
  const [error, setError]       = useState(null)
  const [productoSel, setProductoSel] = useState('')
  const [colorSel, setColorSel]       = useState('')
  // Turno: pre-cargado del turno activo pero editable (puede ser un nuevo período)
  const [turnoNuevo, setTurnoNuevo] = useState(turnoActivo?.Turno || 'M')

  // Pre-seleccionar operario por email (entrante)
  useEffect(() => {
    if (!operarios.length || !usuario?.email) return
    const match = operarios.find(o =>
      (o.Email || '').toLowerCase() === usuario.email.toLowerCase()
    )
    if (match) setValue('Operario', String(match.ID), { shouldValidate: false })
  }, [operarios, usuario?.email])

  // Pre-cargar según el modo:
  // - Relevo: mismo producto (read-only) + mismo color (editable, pre-seleccionado)
  // - Cambio producto: mismo producto (pre-seleccionado pero editable), color vacío
  useEffect(() => {
    if (!turnoActivo || !modoRelevo) return
    if (turnoActivo.Producto) {
      setProductoSel(turnoActivo.Producto)
      setValue('Producto', turnoActivo.Producto, { shouldValidate: false })
    }
    if (turnoActivo.Color) {
      setColorSel(turnoActivo.Color)
      setValue('Color', turnoActivo.Color, { shouldValidate: false })
    }
  }, [turnoActivo?.Producto, turnoActivo?.Color, modoRelevo])

  const onSubmit = async (data) => {
    setEnviando(true)
    setError(null)
    try {
      const token  = await getToken()
      const ahora  = new Date()
      const opObj  = operarios.find(o => String(o.ID) === String(data.Operario))
      const opNombre = opObj ? (opObj.Nombre || opObj.Title || '') : ''
      const prodObj  = productos.find(p => (p.Codigo || '') === data.Producto)
      const kgPorUnidad = prodObj?.KgPorUnidad ?? 0

      // Codigo_Lote: relevo hereda, cambio de producto genera nuevo
      const codigoLote = modoRelevo
        ? (turnoActivo.Codigo_Lote || `${tabletConfig.codigoMaquina}-${data.Producto}-${format(ahora, 'yyyyMMdd-HHmm')}`)
        : `${tabletConfig.codigoMaquina}-${data.Producto}-${format(ahora, 'yyyyMMdd-HHmm')}`

      // Numero_OF: SIEMPRE se hereda (la OF agrupa toda la sesión de la máquina)
      const numeroOF = turnoActivo.Numero_OF || `OF-${tabletConfig.codigoMaquina}-${format(ahora, 'yyyyMMdd-HHmm')}`

      const nuevoRegistroLocal = {
        Maquina: tabletConfig.codigoMaquina,
        Turno: turnoNuevo,
        Operario: opNombre,
        Operario_Apoyo: null,
        Producto: data.Producto,
        Color: data.Color,
        HoraInicio: ahora.toISOString(),
        Fecha: ahora.toISOString(),
        Estado: 'abierto',
        CheckboxVeracidad: false,
        Tipo_Apertura: 'Continuacion',
        Registro_Previo_ID: turnoActivo.spId || null,
        Paradas: '[]',
        Estado_Validacion: 'Verde',
        KgPorUnidadProducto: kgPorUnidad,
        Codigo_Lote: codigoLote,
        Numero_OF: numeroOF,
      }

      let nuevoSpId = null
      if (token && navigator.onLine) {
        const resultado = await createListItem(token, 'Registro_Produccion', {
          Title: tabletConfig.codigoMaquina,
          Turno: turnoNuevo,
          Operario: opNombre,
          Producto: nuevoRegistroLocal.Producto,
          Color: nuevoRegistroLocal.Color,
          HoraInicio: nuevoRegistroLocal.HoraInicio,
          Fecha: nuevoRegistroLocal.Fecha,
          Estado: 'abierto',
          CheckboxVeracidad: false,
          Tipo_Apertura: 'Continuacion',
          Registro_Previo_ID: turnoActivo.spId || null,
          Paradas: '[]',
          Estado_Validacion: 'Verde',
          Codigo_Lote: codigoLote,
          Numero_OF: numeroOF,
        })
        nuevoSpId = resultado.ID
      } else {
        const localId = crypto.randomUUID()
        await encolarOperacion({ tipo: 'create', listName: 'Registro_Produccion', data: nuevoRegistroLocal, localId })
      }

      setModoRelevo(false) // resetear modo relevo después de crear el nuevo registro
      await setTurnoActivo({ ...nuevoRegistroLocal, spId: nuevoSpId, localId: crypto.randomUUID() })
    } catch (err) {
      setError('Error: ' + err.message)
      setEnviando(false)
    }
  }

  return (
    <div style={{ backgroundColor: '#f5f5f5', minHeight: '100vh' }}>
      <Header
        titulo={modoRelevo ? 'Mismo producto' : 'Nuevo producto'}
        color={modoRelevo ? '#37BEEC' : '#004895'}
        pendingCount={pendingCount}
        onLogout={logout}
      />

      <div style={{ padding: '16px', maxWidth: '520px', margin: '0 auto', display: 'flex', flexDirection: 'column', gap: '18px' }}>

        {/* Info del contexto */}
        <div style={{ backgroundColor: modoRelevo ? '#37BEEC' : '#004895', borderRadius: '12px', padding: '12px 16px', color: 'white' }}>
          <p style={{ fontSize: '12px', opacity: 0.8, textTransform: 'uppercase', marginBottom: '4px' }}>
            {modoRelevo ? '🔄 Relevo — Continuación del mismo producto' : '📦 Nuevo producto · Continuación'}
          </p>
          <p style={{ fontSize: '15px', fontWeight: 700 }}>{tabletConfig.nombreMaquina}</p>
          <p style={{ fontSize: '13px', opacity: 0.85, marginTop: '4px' }}>
            {modoRelevo
              ? `Producto en curso: ${turnoActivo?.Producto} · Selecciona el operario entrante`
              : `Producto anterior: ${turnoActivo?.Producto} · Selecciona el nuevo producto`}
          </p>
          {modoRelevo && turnoActivo?.Codigo_Lote && (
            <p style={{ fontSize: '11px', opacity: 0.7, marginTop: '4px', fontFamily: 'monospace' }}>
              🔖 Lote: {turnoActivo.Codigo_Lote}
            </p>
          )}
        </div>

        {/* Selector de turno — editable en caso de cambio de período */}
        <div style={{ backgroundColor: 'white', borderRadius: '12px', padding: '14px 16px', boxShadow: '0 1px 4px rgba(0,0,0,0.08)' }}>
          <label style={{ fontWeight: 600, fontSize: '13px', color: '#333', display: 'block', marginBottom: '10px' }}>
            Turno {turnoNuevo !== turnoActivo?.Turno && <span style={{ color: '#F8A12F', fontSize: '12px' }}> (cambiado)</span>}
          </label>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '8px' }}>
            {[
              { id: 'M', label: 'Mañana' },
              { id: 'T', label: 'Tarde' },
              { id: 'N', label: 'Noche' },
            ].map(t => (
              <button key={t.id} type="button" onClick={() => setTurnoNuevo(t.id)} style={{
                padding: '12px 6px', borderRadius: '8px', border: '2px solid',
                borderColor: turnoNuevo === t.id ? (modoRelevo ? '#37BEEC' : '#004895') : '#ddd',
                backgroundColor: turnoNuevo === t.id ? (modoRelevo ? '#37BEEC' : '#004895') : 'white',
                color: turnoNuevo === t.id ? 'white' : '#333',
                fontWeight: 700, fontSize: '14px', cursor: 'pointer',
              }}>
                {t.label}
              </button>
            ))}
          </div>
        </div>

        <form onSubmit={handleSubmit(onSubmit)} style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>

          {/* Operario */}
          <Field label="Operario *" error={errors.Operario?.message}>
            <select {...register('Operario', { required: 'Selecciona un operario' })} style={inputStyle}>
              <option value="">Seleccionar operario...</option>
              {operarios.map(op => {
                const nombre = op.Nombre || op.Title || ''
                return <option key={op.ID} value={String(op.ID)}>{nombre}</option>
              })}
            </select>
          </Field>

          {/* Producto — bloqueado en relevo, editable en cambio de producto */}
          {modoRelevo ? (
            <div>
              <label style={{ fontWeight: 600, fontSize: '14px', color: '#333', display: 'block', marginBottom: '6px' }}>
                Producto (continuación del turno anterior)
              </label>
              <div style={{ ...inputStyle, backgroundColor: '#f0f4ff', color: '#333', borderColor: '#37BEEC' }}>
                {(() => {
                  const prod = productos.find(p => (p.Codigo || '') === turnoActivo?.Producto || (p.Nombre || p.Title || '') === turnoActivo?.Producto)
                  const nombre = prod ? (prod.Nombre || prod.Title || turnoActivo?.Producto) : turnoActivo?.Producto
                  const codigo = prod?.Codigo || ''
                  return (
                    <div>
                      <strong style={{ fontSize: '15px' }}>{nombre}</strong>
                      {codigo && <span style={{ fontSize: '12px', color: '#888', marginLeft: '8px' }}>({codigo})</span>}
                    </div>
                  )
                })()}
              </div>
              <input type="hidden" {...register('Producto')} />
            </div>
          ) : (
            <Field label="Nuevo producto *" error={errors.Producto?.message}>
              <input type="hidden" {...register('Producto', { required: 'Selecciona un producto' })} />
              <SearchSelect
                opciones={productos.map(p => ({
                  value: p.Codigo || p.Nombre || p.Title || '',
                  label: (p.Nombre || p.Title || '') + (p.Codigo ? ` (${p.Codigo})` : ''),
                }))}
                value={productoSel}
                onChange={v => { setProductoSel(v); setValue('Producto', v, { shouldValidate: true }) }}
                placeholder="Buscar producto..."
              />
            </Field>
          )}

          {/* Color */}
          <Field label={`Color *${modoRelevo ? ' (puede cambiarse)' : ''}`} error={errors.Color?.message}>
            <input type="hidden" {...register('Color', { required: 'Selecciona un color' })} />
            <SearchSelect
              opciones={colores.map(c => ({
                value: c.Nombre || c.Title || '',
                label: c.Nombre || c.Title || '',
              }))}
              value={colorSel}
              onChange={v => { setColorSel(v); setValue('Color', v, { shouldValidate: true }) }}
              placeholder="Buscar color..."
            />
          </Field>


          {error && (
            <div style={{ backgroundColor: '#ffebee', border: '1px solid #f44336', borderRadius: '8px', padding: '12px', color: '#c62828', fontSize: '14px' }}>
              {error}
            </div>
          )}

          <button type="submit" disabled={enviando || cargando} style={{
            backgroundColor: enviando ? '#ccc' : modoRelevo ? '#37BEEC' : '#004895', color: 'white',
            border: 'none', borderRadius: '14px', padding: '18px',
            fontSize: '18px', fontWeight: 700, minHeight: '60px',
            cursor: enviando ? 'not-allowed' : 'pointer', marginTop: '4px',
            boxShadow: enviando ? 'none' : modoRelevo ? '0 4px 16px rgba(55,190,236,0.4)' : '0 4px 16px rgba(0,72,149,0.3)',
          }}>
            {enviando ? '⏳ Iniciando...' : modoRelevo ? '🔄 Iniciar relevo' : '▶ Iniciar nuevo producto'}
          </button>

          <button type="button" onClick={() => setPantalla('bienvenida')} style={{
            backgroundColor: 'transparent', color: '#555',
            border: '2px solid #ccc', borderRadius: '12px',
            padding: '14px', fontSize: '15px', cursor: 'pointer',
          }}>
            ← Volver
          </button>

        </form>
      </div>
    </div>
  )
}
