import { useEffect, useRef, useCallback } from 'react'
import { obtenerPendientes, marcarSincronizado, marcarError } from '../services/indexedDB'
import { createListItem, updateListItem, getSpToken } from '../services/sharepoint'

// Procesa la cola de operaciones pendientes en orden FIFO
export const useSyncQueue = ({ msalInstance, onSyncComplete }) => {
  const syncingRef = useRef(false)

  const procesarCola = useCallback(async () => {
    if (syncingRef.current || !navigator.onLine) return
    const token = await getSpToken(msalInstance)
    if (!token) return

    syncingRef.current = true
    try {
      const pendientes = await obtenerPendientes()
      for (const op of pendientes) {
        try {
          if (op.tipo === 'create') {
            const result = await createListItem(token, op.listName, op.data)
            // Guardar el spId resultante en turno activo si aplica
            if (op.onCreated) op.onCreated(result.ID)
          } else if (op.tipo === 'update') {
            await updateListItem(token, op.listName, op.spId, op.data)
          }
          await marcarSincronizado(op.localId)
        } catch {
          await marcarError(op.localId)
        }
      }
      onSyncComplete?.()
    } finally {
      syncingRef.current = false
    }
  }, [msalInstance, onSyncComplete])

  useEffect(() => {
    const onOnline = () => procesarCola()
    window.addEventListener('online', onOnline)
    // Intentar sync al montar si ya hay conexión
    if (navigator.onLine) procesarCola()
    return () => window.removeEventListener('online', onOnline)
  }, [procesarCola])

  return { sincronizar: procesarCola }
}
