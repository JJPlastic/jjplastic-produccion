import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { PublicClientApplication, EventType } from '@azure/msal-browser'
import { MsalProvider } from '@azure/msal-react'
import { msalConfig } from './config/authConfig'
import './index.css'
import App from './App.jsx'

// Cuando un chunk JS no existe tras un nuevo deploy, Vite falla al importarlo
// dinámicamente. Recargar obtiene el nuevo index.html con los hashes actuales.
window.addEventListener('vite:preloadError', () => {
  window.location.reload()
})

async function bootstrap() {
  const msalInstance = new PublicClientApplication(msalConfig)
  // MSAL v3+ requiere initialize() antes de cualquier operación
  await msalInstance.initialize()

  // Restaurar cuenta activa desde caché (crítico para reload de página)
  if (!msalInstance.getActiveAccount()) {
    const accounts = msalInstance.getAllAccounts()
    if (accounts.length > 0) msalInstance.setActiveAccount(accounts[0])
  }

  // Mantener cuenta activa sincronizada tras login/token refresh
  msalInstance.addEventCallback((event) => {
    if (
      (event.eventType === EventType.LOGIN_SUCCESS ||
        event.eventType === EventType.ACQUIRE_TOKEN_SUCCESS) &&
      event.payload?.account
    ) {
      msalInstance.setActiveAccount(event.payload.account)
    }
  })

  createRoot(document.getElementById('root')).render(
    <StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </StrictMode>
  )
}

bootstrap()
