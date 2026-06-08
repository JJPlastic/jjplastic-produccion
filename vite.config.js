import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { VitePWA } from 'vite-plugin-pwa'
import tailwindcss from '@tailwindcss/vite'

const SP_BASE = '/sites/ProduccionJJPlastic/SiteAssets/app/'

export default defineConfig(({ mode }) => ({
  // En producción los assets apuntan a la ruta de SharePoint
  // En desarrollo apuntan a / (localhost)
  base: '/',
  plugins: [
    react(),
    tailwindcss(),
    VitePWA({
      registerType: 'autoUpdate',
      // Activar Service Worker inmediatamente sin esperar recarga
      injectRegister: 'auto',
      includeAssets: ['favicon.svg', 'pwa-192x192.png', 'pwa-512x512.png'],
      manifest: {
        name: 'JJ Plastic — Producción',
        short_name: 'JJPlastic',
        description: 'Sistema de registro de producción JJ PLASTIC SAC',
        theme_color: '#004895',
        background_color: '#004895',
        display: 'standalone',
        orientation: 'portrait',
        start_url: '/',
        scope: '/',
        lang: 'es',
        icons: [
          { src: 'pwa-192x192.png', sizes: '192x192', type: 'image/png' },
          { src: 'pwa-512x512.png', sizes: '512x512', type: 'image/png', purpose: 'any maskable' },
        ],
      },
      workbox: {
        // Cachear todos los assets del build
        globPatterns: ['**/*.{js,css,html,ico,png,svg,woff2}'],
        // Estrategia Network-First para SharePoint (datos frescos > cache)
        runtimeCaching: [
          {
            urlPattern: /^https:\/\/jjplastic\.sharepoint\.com\/.*/i,
            handler: 'NetworkFirst',
            options: {
              cacheName: 'sharepoint-api',
              networkTimeoutSeconds: 5,
              expiration: { maxEntries: 50, maxAgeSeconds: 60 * 60 * 24 },
              cacheableResponse: { statuses: [0, 200] },
            },
          },
          {
            urlPattern: /^https:\/\/login\.microsoftonline\.com\/.*/i,
            handler: 'NetworkOnly', // MSAL nunca debe cachearse
          },
        ],
        // Forzar activación del SW sin esperar que el usuario cierre todas las tabs
        skipWaiting: true,
        clientsClaim: true,
      },
      devOptions: {
        enabled: false, // No activar SW en dev para evitar interferencias con HMR
      },
    }),
  ],
}))
