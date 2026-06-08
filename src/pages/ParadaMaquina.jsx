import { useState, useEffect, useRef } from 'react'
import { differenceInSeconds, differenceInMinutes } from 'date-fns'
import { useMsal } from '../hooks/useMsal'
import { useApp } from '../context/AppContext'
import { Header } from '../components/Header'
import { LoadingSpinner } from '../components/LoadingSpinner'
import { useMaestros } from '../hooks/useMaestros'
import { updateListItem, uploadAttachment } from '../services/sharepoint'
import { encolarOperacion } from '../services/indexedDB'

const formatCronometro = (secs) => {
  const m = Math.floor(secs / 60)
  const s = secs % 60
  return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`
}

export default function ParadaMaquina() {
  const { getToken, logout } = useMsal()
  const { turnoActivo, actualizarTurnoLocal, pendingCount, setPantalla } = useApp()
  const { motivos, cargando } = useMaestros(getToken)

  const [inicioParada] = useState(() => new Date())
  const [segundos, setSegundos] = useState(0)
  const [motivoSel, setMotivoSel] = useState(null)
  const [descripcion, setDescripcion] = useState('')
  const [foto, setFoto] = useState(null)
  const [fotoNombre, setFotoNombre] = useState('')
  const [guardando, setGuardando] = useState(false)
  const [error, setError] = useState(null)
  const fotoRef = useRef(null)

  // Cronómetro de parada
  useEffect(() => {
    const timer = setInterval(() => {
      setSegundos(differenceInSeconds(new Date(), inicioParada))
    }, 1000)
    return () => clearInterval(timer)
  }, [inicioParada])

  const duracionMinutos = differenceInMinutes(new Date(), inicioParada)
  const fotoRequerida = duracionMinutos >= 30 && !foto

  const handleFotoChange = (e) => {
    const file = e.target.files?.[0]
    if (!file) return
    setFoto(file)
    setFotoNombre(file.name)
  }

  const handleGuardar = async () => {
    setError(null)
    if (!motivoSel) { setError('Selecciona un motivo de parada.'); return }
    if (motivoSel === 'Otro' && !descripcion.trim()) { setError('La descripción es obligatoria cuando el motivo es "Otro".'); return }
    if (fotoRequerida) { setError('La foto es obligatoria para paradas de 30 minutos o más.'); return }

    setGuardando(true)
    const fin = new Date()
    const durSegundos = differenceInSeconds(fin, inicioParada)
    const durMin = parseFloat((durSegundos / 60).toFixed(2)) // decimal, ej: 1.5 min
    const nuevaParada = {
      id: crypto.randomUUID(),
      motivo: motivoSel,
      descripcion: descripcion.trim(),
      timestamp_inicio: inicioParada.toISOString(),
      timestamp_fin: fin.toISOString(),
      duracion_segundos: durSegundos,
      duracion_minutos:  durMin,
      foto_url: null,
    }

    try {
      // Subir foto si existe y hay spId
      if (foto && turnoActivo.spId) {
        const token = await getToken()
        if (token) {
          const nombre = `parada_${turnoActivo.spId}_${Date.now()}.jpg`
          const url = await uploadAttachment(token, 'Registro_Produccion', turnoActivo.spId, nombre, foto)
          nuevaParada.foto_url = url
        }
      }

      // Actualizar array de paradas
      const paradasActuales = (() => {
        try { return JSON.parse(turnoActivo.Paradas || '[]') }
        catch { return [] }
      })()
      const paradasNuevas = [...paradasActuales, nuevaParada]
      const payload = { Paradas: JSON.stringify(paradasNuevas) }

      if (turnoActivo.spId && navigator.onLine) {
        const token = await getToken()
        if (token) await updateListItem(token, 'Registro_Produccion', turnoActivo.spId, payload)
      } else {
        await encolarOperacion({
          tipo: 'update',
          listName: 'Registro_Produccion',
          spId: turnoActivo.spId,
          data: payload,
        })
      }

      // Actualizar estado local inmediatamente
      await actualizarTurnoLocal({ Paradas: JSON.stringify(paradasNuevas) })
      setPantalla('turno-activo')
    } catch (err) {
      setError('Error al guardar: ' + err.message)
    } finally {
      setGuardando(false)
    }
  }

  if (cargando) return <LoadingSpinner mensaje="Cargando motivos..." />

  return (
    <div style={{ backgroundColor: '#f0f2f5', minHeight: '100vh' }}>
      <Header titulo="Parada de máquina" pendingCount={pendingCount} onLogout={logout} color="#c62828" />

      <div style={{ padding: '12px', maxWidth: '480px', margin: '0 auto', display: 'flex', flexDirection: 'column', gap: '10px' }}>

        {/* Cronómetro */}
        <div style={{ backgroundColor: '#c62828', borderRadius: '12px', padding: '16px', color: 'white', textAlign: 'center' }}>
          <p style={{ fontSize: '10px', opacity: 0.75, textTransform: 'uppercase', letterSpacing: '0.08em', margin: 0 }}>Tiempo de parada</p>
          <p style={{ fontSize: '48px', fontWeight: 800, fontVariantNumeric: 'tabular-nums', lineHeight: 1.1, margin: '4px 0 0' }}>
            {formatCronometro(segundos)}
          </p>
          {duracionMinutos >= 30 && (
            <p style={{ fontSize: '12px', opacity: 0.9, margin: '6px 0 0' }}>📷 Foto obligatoria</p>
          )}
        </div>

        {/* Motivos — 3 columnas */}
        <div>
          <p style={{ fontWeight: 700, fontSize: '12px', color: '#555', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: '7px' }}>
            Motivo *
          </p>
          {motivos.length === 0 ? (
            <p style={{ color: '#888', fontSize: '13px' }}>Sin motivos disponibles offline.</p>
          ) : (
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '6px' }}>
              {motivos.map((m) => {
                const motivo = m.Motivo || m.Title || '(sin nombre)'
                const sel = motivoSel === motivo
                return (
                  <button key={m.ID} type="button" onClick={() => setMotivoSel(motivo)} style={{
                    padding: '10px 6px', borderRadius: '8px', border: '2px solid',
                    borderColor: sel ? '#c62828' : '#ddd',
                    backgroundColor: sel ? '#c62828' : 'white',
                    color: sel ? 'white' : '#333',
                    fontWeight: sel ? 700 : 500,
                    fontSize: '12px', minHeight: '52px', cursor: 'pointer',
                    textAlign: 'center', lineHeight: 1.3,
                  }}>
                    {motivo}
                  </button>
                )
              })}
            </div>
          )}
        </div>

        {/* Descripción */}
        <div>
          <label style={{ fontWeight: 700, fontSize: '12px', color: '#555', textTransform: 'uppercase', letterSpacing: '0.05em', display: 'block', marginBottom: '5px' }}>
            Descripción {motivoSel === 'Otro' ? '*' : '(opcional)'}
          </label>
          <textarea value={descripcion} onChange={e => setDescripcion(e.target.value)} rows={2}
            placeholder="Describe brevemente el motivo..."
            style={{ width: '100%', padding: '10px', borderRadius: '8px', border: '2px solid #ddd', fontSize: '14px', resize: 'none', fontFamily: 'inherit', backgroundColor: 'white', color: '#1a1a1a', boxSizing: 'border-box' }}
          />
        </div>

        {/* Foto */}
        <button type="button" onClick={() => fotoRef.current?.click()} style={{
          width: '100%', padding: '12px', borderRadius: '8px',
          border: `2px dashed ${foto ? '#4caf50' : fotoRequerida ? '#d32f2f' : '#ccc'}`,
          backgroundColor: foto ? '#f1f8e9' : 'white',
          color: foto ? '#2e7d32' : fotoRequerida ? '#d32f2f' : '#888',
          fontSize: '14px', fontWeight: 600, cursor: 'pointer',
        }}>
          {foto ? `✓ ${fotoNombre}` : `📷 Foto ${duracionMinutos >= 30 ? '(obligatoria)' : '(opcional)'}`}
        </button>
        <input ref={fotoRef} type="file" accept="image/*" capture="environment" style={{ display: 'none' }} onChange={handleFotoChange} />

        {error && (
          <div style={{ backgroundColor: '#ffebee', border: '1px solid #f44336', borderRadius: '8px', padding: '10px 12px', color: '#c62828', fontSize: '13px', fontWeight: 600 }}>
            {error}
          </div>
        )}

        {/* Botones */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', gap: '8px', marginTop: '2px' }}>
          <button type="button" onClick={() => setPantalla('turno-activo')} style={{
            padding: '14px', borderRadius: '10px', border: '2px solid #ccc',
            backgroundColor: 'white', color: '#555', fontSize: '14px', fontWeight: 600, cursor: 'pointer',
          }}>Cancelar</button>
          <button type="button" onClick={handleGuardar} disabled={guardando} style={{
            padding: '14px', borderRadius: '10px', border: 'none',
            backgroundColor: guardando ? '#ccc' : '#c62828', color: 'white',
            fontSize: '15px', fontWeight: 700, cursor: guardando ? 'not-allowed' : 'pointer',
          }}>
            {guardando ? '⏳ Guardando...' : '▶ Registrar y reanudar'}
          </button>
        </div>

      </div>
    </div>
  )
}
