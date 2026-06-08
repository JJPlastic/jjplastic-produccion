import { useEffect, useRef } from 'react'

// Mantiene la pantalla encendida mientras la app está activa (crítico en tablets de fábrica)
export const useWakeLock = (enabled = true) => {
  const lockRef = useRef(null)

  const acquire = async () => {
    if (!('wakeLock' in navigator) || !enabled) return
    try {
      lockRef.current = await navigator.wakeLock.request('screen')
    } catch {
      // WakeLock no disponible o denegado — no es crítico
    }
  }

  const release = () => {
    lockRef.current?.release()
    lockRef.current = null
  }

  useEffect(() => {
    if (!enabled) return
    acquire()
    // Re-adquirir cuando la página vuelve a ser visible (tab switch, bloqueo de pantalla)
    const onVisibilityChange = () => {
      if (document.visibilityState === 'visible') acquire()
    }
    document.addEventListener('visibilitychange', onVisibilityChange)
    return () => {
      release()
      document.removeEventListener('visibilitychange', onVisibilityChange)
    }
  }, [enabled])
}
