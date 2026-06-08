// Genera pwa-192x192.png y pwa-512x512.png sin dependencias externas
import { createWriteStream } from 'fs'
import { deflateSync } from 'zlib'
import { mkdir } from 'fs/promises'

const NAVY  = [0x00, 0x48, 0x95]
const ORANGE = [0xF8, 0xA1, 0x2F]
const WHITE  = [0xFF, 0xFF, 0xFF]

function hexToRgb(hex) {
  const n = parseInt(hex.replace('#',''), 16)
  return [(n >> 16) & 0xff, (n >> 8) & 0xff, n & 0xff]
}

// Escribe un chunk PNG válido
function pngChunk(type, data) {
  const typeBuf = Buffer.from(type, 'ascii')
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length)
  const crcBuf = Buffer.allocUnsafe(4)
  // CRC32 del tipo + datos
  const crcData = Buffer.concat([typeBuf, data])
  crcBuf.writeUInt32BE(crc32(crcData))
  return Buffer.concat([len, typeBuf, data, crcBuf])
}

// CRC32
function crc32(buf) {
  let table = []
  for (let i = 0; i < 256; i++) {
    let c = i
    for (let j = 0; j < 8; j++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1
    table[i] = c
  }
  let crc = 0xFFFFFFFF
  for (let b of buf) crc = table[(crc ^ b) & 0xff] ^ (crc >>> 8)
  return (crc ^ 0xFFFFFFFF) >>> 0
}

// Dibuja un icono simple: fondo navy, cuadrado redondeado orange, letras JJ
function drawIcon(size) {
  const pixels = []
  const cx = size / 2, cy = size / 2
  const pad = size * 0.15
  const r = size * 0.2 // radio del rectángulo interno

  for (let y = 0; y < size; y++) {
    const row = []
    for (let x = 0; x < size; x++) {
      // Fondo: círculo navy con esquinas redondeadas
      const distCorner = Math.min(
        Math.hypot(x - size*0.18, y - size*0.18),
        Math.hypot(x - size*0.82, y - size*0.18),
        Math.hypot(x - size*0.18, y - size*0.82),
        Math.hypot(x - size*0.82, y - size*0.82),
      )
      const inRoundedRect = x >= size*0.18 && x <= size*0.82 && y >= size*0.18 && y <= size*0.82
      const inCornerRadius = distCorner <= size * 0.18

      // Rectángulo interior naranja
      const ix1 = size * 0.25, ix2 = size * 0.75
      const iy1 = size * 0.25, iy2 = size * 0.75
      const inOrange = x >= ix1 && x <= ix2 && y >= iy1 && y <= iy2

      // Letras "JJ" simplificadas como píxeles blancos
      const relX = (x - cx) / size
      const relY = (y - cy) / size

      // J izquierda: línea vertical -0.15 a -0.05 de -0.2 a 0.2, hook abajo
      const jLeft = (relX >= -0.18 && relX <= -0.05) && (relY >= -0.22 && relY <= 0.22) &&
        (Math.abs(relX + 0.115) <= 0.065)
      const jLeftHook = relY >= 0.08 && relY <= 0.22 && relX >= -0.22 && relX <= -0.05 &&
        Math.hypot(relX + 0.05, relY - 0.08) <= 0.14 && relX < -0.05

      // J derecha: línea vertical 0.05 a 0.18 de -0.2 a 0.2, hook abajo
      const jRight = (relX >= 0.04 && relX <= 0.18) && (relY >= -0.22 && relY <= 0.22) &&
        (Math.abs(relX - 0.11) <= 0.065)
      const jRightHook = relY >= 0.08 && relY <= 0.22 && relX >= -0.02 && relX <= 0.18 &&
        Math.hypot(relX - 0.04, relY - 0.08) <= 0.14 && relX < 0.04

      const isLetter = (jLeft || jRight) && inOrange

      if (isLetter) {
        row.push(...WHITE, 0xff)
      } else if (inOrange && (inRoundedRect || !inCornerRadius)) {
        row.push(...ORANGE, 0xff)
      } else if (inRoundedRect || inCornerRadius) {
        row.push(...NAVY, 0xff)
      } else {
        row.push(0, 0, 0, 0) // transparente
      }
    }
    pixels.push(row)
  }
  return pixels
}

function buildPng(size) {
  const pixels = drawIcon(size)
  const sig = Buffer.from([137,80,78,71,13,10,26,10])

  // IHDR
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(size, 0)
  ihdr.writeUInt32BE(size, 4)
  ihdr[8] = 8   // bit depth
  ihdr[9] = 6   // RGBA
  ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0

  // Raw image data (filtro 0 por fila)
  const raw = []
  for (const row of pixels) {
    raw.push(0, ...row)
  }
  const compressed = deflateSync(Buffer.from(raw), { level: 6 })

  return Buffer.concat([
    sig,
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', compressed),
    pngChunk('IEND', Buffer.alloc(0)),
  ])
}

await mkdir('public', { recursive: true })

for (const size of [192, 512]) {
  const png = buildPng(size)
  const path = `public/pwa-${size}x${size}.png`
  const ws = createWriteStream(path)
  ws.write(png)
  ws.end()
  console.log(`✓ ${path} (${png.length} bytes)`)
}
