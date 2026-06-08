import { useMsal as useMsalOriginal } from '@azure/msal-react'
import { loginRequest, SHAREPOINT_SCOPE } from '../config/authConfig'

export const useMsal = () => {
  const { instance, accounts } = useMsalOriginal()

  const login = () => instance.loginRedirect(loginRequest)

  const logout = () =>
    instance.logoutRedirect({ postLogoutRedirectUri: window.location.origin })

  // Obtiene token específico para SharePoint REST API
  const getToken = async () => {
    const account = instance.getActiveAccount() || accounts[0]
    if (!account) return null
    try {
      const res = await instance.acquireTokenSilent({
        scopes: [SHAREPOINT_SCOPE],
        account,
      })
      return res.accessToken
    } catch {
      // Silent falló (expirado, consent requerido) → redirect
      await instance.acquireTokenRedirect({ scopes: [SHAREPOINT_SCOPE], account })
      return null
    }
  }

  const usuario = accounts[0]
    ? { nombre: accounts[0].name, email: accounts[0].username }
    : null

  return {
    login,
    logout,
    getToken,
    usuario,
    msalInstance: instance,
    isAuthenticated: accounts.length > 0,
  }
}
