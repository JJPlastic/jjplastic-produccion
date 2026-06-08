export const msalConfig = {
  auth: {
    clientId: '1f17f898-96f3-4730-a654-6b919da9137a',
    authority: 'https://login.microsoftonline.com/9ec268d0-652e-457f-81e7-3f71408dea4b',
    // BASE_URL en dev = '/'  →  http://localhost:5174/
    // BASE_URL en prod = '/sites/ProduccionJJPlastic/SiteAssets/app/'
    redirectUri: window.location.origin + import.meta.env.BASE_URL,
    postLogoutRedirectUri: window.location.origin + import.meta.env.BASE_URL,
    navigateToLoginRequestUrl: true,
  },
  cache: {
    cacheLocation: 'localStorage',
    // storeAuthStateInCookie: true es crítico para Safari/iOS en modo standalone PWA
    storeAuthStateInCookie: true,
  },
  system: {
    allowNativeBroker: false,
    loggerOptions: {
      logLevel: import.meta.env.DEV ? 3 : 0, // Verbose en dev, silent en prod
    }
  }
}

// Scope de SharePoint como constante exportable para usarlo en toda la app
export const SHAREPOINT_SCOPE = 'https://jjplastic.sharepoint.com/AllSites.Write'

// Login solicita ambos recursos — MSAL los separa internamente en tokens distintos
export const loginRequest = {
  scopes: ['User.Read', SHAREPOINT_SCOPE]
}

// Para acquireTokenSilent solo se puede pedir UN recurso a la vez
export const spTokenRequest = {
  scopes: [SHAREPOINT_SCOPE]
}

export const sharePointConfig = {
  siteUrl: 'https://jjplastic.sharepoint.com/sites/ProduccionJJPlastic',
  apiUrl: 'https://jjplastic.sharepoint.com/sites/ProduccionJJPlastic/_api/web/lists',
}
