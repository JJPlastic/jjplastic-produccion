import { sharePointConfig, SHAREPOINT_SCOPE } from '../config/authConfig'

export const getSpToken = async (msalInstance) => {
  const account = msalInstance.getActiveAccount()
  if (!account) return null
  try {
    const res = await msalInstance.acquireTokenSilent({ scopes: [SHAREPOINT_SCOPE], account })
    return res.accessToken
  } catch {
    await msalInstance.acquireTokenRedirect({ scopes: [SHAREPOINT_SCOPE], account })
    return null
  }
}

// odata=nometadata: no requiere __metadata en el body — evita error de tipo inválido
const baseHeaders = (token) => ({
  Authorization: `Bearer ${token}`,
  Accept: 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata',
})

// ensureuser: agrega al sitio si no existe y devuelve el ID SP
// Más confiable que siteusers/?$filter pues no requiere visita previa al sitio
export const resolveUserId = async (token, email) => {
  try {
    const res = await fetch(`${sharePointConfig.siteUrl}/_api/web/ensureuser`, {
      method: 'POST',
      headers: baseHeaders(token),
      body: JSON.stringify({ logonName: `i:0#.f|membership|${email}` }),
    })
    if (!res.ok) return null
    const data = await res.json()
    return data.Id ?? data.d?.Id ?? null
  } catch {
    return null
  }
}

// ─── Lecturas ─────────────────────────────────────────────────────────────────

export const getListItems = async (token, listName, { filter, select, orderby, top } = {}) => {
  const params = new URLSearchParams()
  if (filter)  params.set('$filter', filter)
  if (select)  params.set('$select', select)
  if (orderby) params.set('$orderby', orderby)
  if (top)     params.set('$top', String(top))
  const query = params.toString() ? `?${params}` : ''
  const url = `${sharePointConfig.apiUrl}/getbytitle('${listName}')/items${query}`
  const res = await fetch(url, { headers: baseHeaders(token) })
  if (!res.ok) {
    const err = await res.text()
    throw new Error(`getListItems(${listName}): ${res.status} — ${err}`)
  }
  const data = await res.json()
  return data.value ?? data.d?.results ?? []
}

export const getListItem = async (token, listName, itemId) => {
  const url = `${sharePointConfig.apiUrl}/getbytitle('${listName}')/items(${itemId})`
  const res = await fetch(url, { headers: baseHeaders(token) })
  if (!res.ok) throw new Error(`getListItem(${listName}, ${itemId}): ${res.status}`)
  const data = await res.json()
  return data
}

// ─── Escrituras ───────────────────────────────────────────────────────────────

export const createListItem = async (token, listName, item) => {
  const url = `${sharePointConfig.apiUrl}/getbytitle('${listName}')/items`
  const res = await fetch(url, {
    method: 'POST',
    headers: baseHeaders(token),
    body: JSON.stringify(item),
  })
  if (!res.ok) {
    const err = await res.text()
    throw new Error(`createListItem(${listName}): ${res.status} — ${err}`)
  }
  return await res.json()
}

export const updateListItem = async (token, listName, itemId, item) => {
  const url = `${sharePointConfig.apiUrl}/getbytitle('${listName}')/items(${itemId})`
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      ...baseHeaders(token),
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*',
    },
    body: JSON.stringify(item),
  })
  if (!res.ok) {
    const err = await res.text()
    throw new Error(`updateListItem(${listName}, ${itemId}): ${res.status} — ${err}`)
  }
  return true
}

// ─── Órdenes de Fabricación activas para una máquina ─────────────────────────
// Busca en Kardex_MP (OFs con MP entregada) y Registro_Produccion (OFs en producción)
// Combina ambas fuentes para mostrar el estado completo de cada OF
export const getOFsActivas = async (token, maquina) => {
  try {
    // Carga registros filtrando por máquina cliente-side (OData Title filter es unreliable)
    const [rKardex, rRegistros] = await Promise.allSettled([
      getListItems(token, 'Kardex_MP', { top: 200 }),
      getListItems(token, 'Registro_Produccion', { top: 200 }),
    ])

    const mapa = {}

    // ── Fuente 1: Kardex_MP con Numero_OF ────────────────────────────────────
    if (rKardex.status === 'fulfilled') {
      rKardex.value
        .filter(k => {
          const maqK = (k.Title || '').trim()
          return maqK === maquina && k.Numero_OF && k.Numero_OF.trim() !== ''
        })
        .forEach(k => {
          const of = k.Numero_OF.trim()
          if (!mapa[of]) mapa[of] = {
            of, maquina,
            mp: [], insumos: [],
            productos: new Set(),
            tieneKardex: false,
            tieneRegistro: false,
          }
          mapa[of].tieneKardex = true
          mapa[of].mp.push({
            insumo: k.Insumo || '—',
            kg: parseFloat(k.KgEntregados) || 0,
            kgDev: parseFloat(k.KgDevueltos) || 0,
            producto: k.Producto || '',
          })
          if (k.Insumo && !mapa[of].insumos.includes(k.Insumo)) {
            mapa[of].insumos.push(k.Insumo)
          }
          if (k.Producto) mapa[of].productos.add(k.Producto)
        })
    }

    // ── Fuente 2: Registro_Produccion con Numero_OF (todos los estados) ──────
    if (rRegistros.status === 'fulfilled') {
      rRegistros.value
        .filter(r => {
          const maqR = (r.Title || r.Maquina || '').trim()
          return maqR === maquina && r.Numero_OF && r.Numero_OF.trim() !== ''
        })
        .forEach(r => {
          const of = r.Numero_OF.trim()
          if (!mapa[of]) mapa[of] = {
            of, maquina,
            mp: [], insumos: [],
            productos: new Set(),
            tieneKardex: false,
            tieneRegistro: false,
            tieneCerrado: false,
          }
          if (r.Estado === 'abierto' || r.Estado === 'transferido') {
            mapa[of].tieneRegistro = true
            if (r.Producto) mapa[of].productos.add(r.Producto)
            if (r.Operario) mapa[of].ultimoOperario = r.Operario
          }
          if (r.Estado === 'cerrado') {
            mapa[of].tieneCerrado = true
          }
        })
    }

    // Filtrar OFs cerradas: solo mostrar si hay producción abierta/transferida
    // O si PCP registró MP pero aún no hay ningún Registro (operario no ha iniciado)
    const ofsActivas = Object.values(mapa).filter(o => {
      if (o.tieneRegistro) return true          // producción en curso → siempre mostrar
      if (o.tieneCerrado && !o.tieneRegistro) return false  // turno cerrado → no mostrar
      if (o.tieneKardex && !o.tieneCerrado) return true     // PCP entregó MP, operario pendiente
      return false
    })

    return ofsActivas.map(o => {
      // Consolidar entradas del mismo insumo+producto
      const mpConsolidado = {}
      o.mp.forEach(m => {
        const key = `${m.producto || ''}__${(m.insumo || '').trim().toLowerCase()}`
        if (!mpConsolidado[key]) mpConsolidado[key] = { insumo: m.insumo, kg: 0, kgDev: 0, producto: m.producto || '' }
        mpConsolidado[key].kg    += m.kg    || 0
        mpConsolidado[key].kgDev += m.kgDev || 0
      })
      const mpFinal = Object.values(mpConsolidado)

      // Agrupar por producto → { producto: [{insumo, kg, kgDev}] }
      const mpPorProducto = {}
      mpFinal.forEach(m => {
        const prod = m.producto || ''
        if (!mpPorProducto[prod]) mpPorProducto[prod] = []
        mpPorProducto[prod].push(m)
      })

      return {
        ...o,
        mp: mpFinal,
        mpPorProducto,
        productos: [...o.productos],
        totalKgMP: mpFinal.reduce((s, m) => s + (m.kg - m.kgDev), 0),
        resumenMP: mpFinal.map(m =>
          `${m.insumo}: ${(m.kg - m.kgDev).toFixed(2)} kg`
        ).join(' · '),
      }
    })
  } catch (err) {
    console.error('getOFsActivas error:', err)
    return []
  }
}

// ─── Adjuntos (fotos) ─────────────────────────────────────────────────────────

export const uploadAttachment = async (token, listName, itemId, filename, fileBlob) => {
  const url = `${sharePointConfig.siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(filename)}')`
  const res = await fetch(url, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=nometadata' },
    body: await fileBlob.arrayBuffer(),
  })
  if (!res.ok) {
    const err = await res.text()
    throw new Error(`uploadAttachment: ${res.status} — ${err}`)
  }
  const data = await res.json()
  return data.ServerRelativeUrl ?? data.d?.ServerRelativeUrl
}
