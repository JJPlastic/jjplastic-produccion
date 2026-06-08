import { useState, useRef, useEffect } from 'react'

/**
 * Campo con búsqueda y sugerencias en tiempo real.
 * Reemplaza <select> para listas largas.
 */
export const SearchSelect = ({
  opciones = [],       // [{ value, label }]
  value,
  onChange,
  placeholder = 'Buscar...',
  style = {},
  disabled = false,
}) => {
  const [texto, setTexto]         = useState('')
  const [abierto, setAbierto]     = useState(false)
  const [resaltado, setResaltado] = useState(-1)
  const inputRef                  = useRef(null)
  const listaRef                  = useRef(null)
  const contenedorRef             = useRef(null)

  const labelActual = opciones.find(o => o.value === value)?.label || ''

  const filtradas = texto.trim()
    ? opciones.filter(o => o.label.toLowerCase().includes(texto.toLowerCase())).slice(0, 40)
    : opciones.slice(0, 40)

  const seleccionar = (opcion) => {
    onChange(opcion.value)
    setTexto('')
    setAbierto(false)
    setResaltado(-1)
    inputRef.current?.blur()
  }

  const handleKey = (e) => {
    if (!abierto) { setAbierto(true); return }
    if (e.key === 'ArrowDown') {
      e.preventDefault()
      setResaltado(r => Math.min(r + 1, filtradas.length - 1))
    } else if (e.key === 'ArrowUp') {
      e.preventDefault()
      setResaltado(r => Math.max(r - 1, 0))
    } else if (e.key === 'Enter' && resaltado >= 0) {
      e.preventDefault()
      seleccionar(filtradas[resaltado])
    } else if (e.key === 'Escape') {
      setAbierto(false)
      setTexto('')
    }
  }

  // Cerrar al clic afuera
  useEffect(() => {
    const fn = (e) => {
      if (!contenedorRef.current?.contains(e.target)) {
        setAbierto(false)
        setTexto('')
      }
    }
    document.addEventListener('mousedown', fn)
    return () => document.removeEventListener('mousedown', fn)
  }, [])

  // Scroll al resaltado
  useEffect(() => {
    if (resaltado >= 0 && listaRef.current) {
      listaRef.current.children[resaltado]?.scrollIntoView({ block: 'nearest' })
    }
  }, [resaltado])

  return (
    <div ref={contenedorRef} style={{ position: 'relative', ...style }}>
      <div style={{ position: 'relative' }}>
        <input
          ref={inputRef}
          type="text"
          disabled={disabled}
          placeholder={value ? '' : placeholder}
          value={abierto ? texto : ''}
          onChange={e => { setTexto(e.target.value); setAbierto(true); setResaltado(-1) }}
          onFocus={() => setAbierto(true)}
          onKeyDown={handleKey}
          style={{
            width: '100%',
            padding: '14px 40px 14px 12px',
            borderRadius: '10px',
            border: `2px solid ${abierto ? '#004895' : '#ddd'}`,
            fontSize: '16px',
            backgroundColor: 'white',
            color: '#1a1a1a',
            boxSizing: 'border-box',
            outline: 'none',
          }}
        />

        {/* Etiqueta del valor seleccionado (cuando no se está buscando) */}
        {!abierto && value && (
          <div style={{
            position: 'absolute', inset: 0,
            padding: '14px 40px 14px 12px',
            fontSize: '16px', color: '#1a1a1a',
            pointerEvents: 'none',
            display: 'flex', alignItems: 'center',
            overflow: 'hidden', whiteSpace: 'nowrap',
          }}>
            {labelActual}
          </div>
        )}

        {/* Botón limpiar */}
        {value && !abierto && (
          <button type="button"
            onClick={(e) => { e.stopPropagation(); onChange('') }}
            style={{
              position: 'absolute', right: '28px', top: '50%',
              transform: 'translateY(-50%)',
              background: 'none', border: 'none',
              color: '#999', cursor: 'pointer',
              fontSize: '18px', lineHeight: 1, padding: '0 4px',
            }}>×</button>
        )}

        {/* Flecha */}
        <div style={{
          position: 'absolute', right: '10px', top: '50%',
          transform: 'translateY(-50%)',
          color: '#666', pointerEvents: 'none', fontSize: '11px',
        }}>
          {abierto ? '▲' : '▼'}
        </div>
      </div>

      {/* Lista desplegable */}
      {abierto && filtradas.length > 0 && (
        <div ref={listaRef} style={{
          position: 'absolute', top: '100%', left: 0, right: 0,
          zIndex: 1000,
          backgroundColor: 'white',
          border: '2px solid #004895',
          borderTop: 'none',
          borderRadius: '0 0 10px 10px',
          maxHeight: '240px', overflowY: 'auto',
          boxShadow: '0 6px 16px rgba(0,0,0,0.15)',
        }}>
          {filtradas.map((op, idx) => {
            const lbl = op.label
            const q = texto.toLowerCase()
            const i = q ? lbl.toLowerCase().indexOf(q) : -1
            return (
              <div key={`${op.value}-${idx}`}
                onMouseDown={(e) => { e.preventDefault(); seleccionar(op) }}
                style={{
                  padding: '12px 14px',
                  cursor: 'pointer',
                  fontSize: '15px',
                  color: '#1a1a1a',
                  backgroundColor:
                    idx === resaltado ? '#e8f0fe'
                    : op.value === value ? '#f0f4ff'
                    : 'white',
                  borderBottom: '1px solid #f0f0f0',
                  fontWeight: op.value === value ? 700 : 400,
                  display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                }}>
                <span>
                  {i >= 0 ? (
                    <>
                      {lbl.slice(0, i)}
                      <mark style={{ backgroundColor: '#fff176', borderRadius: '2px', padding: '0 1px', color: '#1a1a1a' }}>
                        {lbl.slice(i, i + q.length)}
                      </mark>
                      {lbl.slice(i + q.length)}
                    </>
                  ) : lbl}
                </span>
                {op.value === value && <span style={{ color: '#004895', fontSize: '13px' }}>✓</span>}
              </div>
            )
          })}
        </div>
      )}

      {/* Sin resultados */}
      {abierto && texto.trim() && filtradas.length === 0 && (
        <div style={{
          position: 'absolute', top: '100%', left: 0, right: 0,
          zIndex: 1000,
          backgroundColor: 'white',
          border: '2px solid #ddd', borderTop: 'none',
          borderRadius: '0 0 10px 10px',
          padding: '12px 14px',
          fontSize: '14px', color: '#888',
        }}>
          Sin resultados para "<strong>{texto}</strong>"
        </div>
      )}
    </div>
  )
}
