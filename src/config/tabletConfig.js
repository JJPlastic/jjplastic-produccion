const STORAGE_KEY         = 'jjplastic_tablet_config'
const STORAGE_KEY_TEMP    = 'jjplastic_tablet_temp'   // solo sesión (sessionStorage)
const TOKEN_KEY           = 'jjplastic_tablet_token'

export const MAQUINAS_DISPONIBLES = [
  { codigoMaquina: 'MAQ01', nombreMaquina: 'GII MA 3600' },
  { codigoMaquina: 'MAQ02', nombreMaquina: 'Haitian 128T' },
  { codigoMaquina: 'MAQ03', nombreMaquina: 'Haitian 130T' },
  { codigoMaquina: 'MAQ04', nombreMaquina: 'Haitian 168T' },
  { codigoMaquina: 'MAQ05', nombreMaquina: 'Haitian 200T' },
  { codigoMaquina: 'MAQ06', nombreMaquina: 'Haitian 360T' },
]

// Token único por tablet — se genera una vez y persiste
export const getTabletToken = () => {
  let token = localStorage.getItem(TOKEN_KEY)
  if (!token) {
    token = `tablet_${Date.now()}_${Math.random().toString(36).slice(2,8)}`
    localStorage.setItem(TOKEN_KEY, token)
  }
  return token
}

const cargarConfig = () => {
  try {
    const guardado = localStorage.getItem(STORAGE_KEY)
    if (guardado) return JSON.parse(guardado)
  // eslint-disable-next-line no-empty
  } catch (_e) { }
  return null
}

export const guardarTabletConfig = (codigoMaquina, nombreMaquina) => {
  const config = { codigoMaquina, nombreMaquina }
  localStorage.setItem(STORAGE_KEY, JSON.stringify(config))
  sessionStorage.removeItem(STORAGE_KEY_TEMP) // limpiar temporal si se hace permanente
  return config
}

export const guardarTabletConfigTemporal = (codigoMaquina, nombreMaquina) => {
  const config = { codigoMaquina, nombreMaquina }
  sessionStorage.setItem(STORAGE_KEY_TEMP, JSON.stringify(config))
  return config
}

export const limpiarTabletConfig = () => {
  localStorage.removeItem(STORAGE_KEY)
  sessionStorage.removeItem(STORAGE_KEY_TEMP)
}

export const tabletConfigGuardada = () => {
  // Guardada permanente O solo esta sesión
  return cargarConfig() !== null || sessionStorage.getItem(STORAGE_KEY_TEMP) !== null
}

// Exportación principal — usa permanente, sino temporal, sino vacío
const cargarConfigActiva = () => {
  const perm = cargarConfig()
  if (perm) return perm
  try {
    const temp = sessionStorage.getItem(STORAGE_KEY_TEMP)
    if (temp) return JSON.parse(temp)
  // eslint-disable-next-line no-empty
  } catch (_e) { }
  return null
}

const config = cargarConfigActiva()
export const tabletConfig = config ?? { codigoMaquina: '', nombreMaquina: '' }
