// import express from 'express'
// import ExcelJS from 'exceljs'
// import fs from 'fs'
// import path from 'path'
// import http from 'http'
// import https from 'https'
// import os from 'os'
// import { fileURLToPath } from 'url'

// const __dirname = path.dirname(fileURLToPath(import.meta.url))
// const PORT = process.env.PORT || 3000
// const FILE = path.join(__dirname, 'data', 'visitors.xlsx')
// const SHEET = 'Visitors'
// const COLUMNS = [
//   { header: 'Sr No', key: 'sr', width: 8 },
//   { header: 'Registration ID', key: 'registrationId', width: 24 },
//   { header: 'Name', key: 'name', width: 26 },
//   { header: 'No. of Visitors', key: 'visitorCount', width: 16 },
//   { header: 'Person to Meet', key: 'personToMeet', width: 24 },
//   { header: 'Visit Date & Time', key: 'visitDateTime', width: 22 },
//   { header: 'Scanned At', key: 'scannedAt', width: 22 },
// ]

// const app = express()
// app.use(express.json({ limit: '100kb' }))
// app.use(express.static(path.join(__dirname, 'public')))
// app.use('/lib/html5-qrcode.min.js', express.static(path.join(__dirname, 'node_modules/html5-qrcode/html5-qrcode.min.js')))

// // Serialize writes so two quick scans never corrupt the file
// let queue = Promise.resolve()
// const enqueue = (fn) => (queue = queue.then(fn, fn))

// const fmtIST = (d) => new Date(d).toLocaleString('en-IN', { timeZone: 'Asia/Kolkata', dateStyle: 'medium', timeStyle: 'medium' })

// async function openBook() {
//   const wb = new ExcelJS.Workbook()
//   if (fs.existsSync(FILE)) {
//     await wb.xlsx.readFile(FILE)
//     return { wb, ws: wb.getWorksheet(SHEET) }
//   }
//   const ws = wb.addWorksheet(SHEET, { views: [{ state: 'frozen', ySplit: 1 }] })
//   ws.columns = COLUMNS
//   const h = ws.getRow(1)
//   h.font = { bold: true, color: { argb: 'FFFFFFFF' } }
//   h.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F7A4D' } }
//   h.alignment = { vertical: 'middle' }
//   return { wb, ws }
// }

// function parsePayload(raw) {
//   let d
//   try { d = typeof raw === 'string' ? JSON.parse(raw) : raw } catch { return null }
//   if (!d || typeof d !== 'object' || !d.registrationId) return null
//   const s = (v) => (v == null ? '' : String(v).slice(0, 200))
//   return {
//     registrationId: s(d.registrationId).trim(),
//     name: s(d.name),
//     visitorCount: Number(d.visitorCount) || 1,
//     personToMeet: s(d.personToMeet),
//     visitDateTime: s(d.visitDateTime).replace('T', ' '),
//   }
// }

// app.post('/api/scan', (req, res) => {
//   const v = parsePayload(req.body?.qr)
//   if (!v) return res.status(400).json({ ok: false, error: 'Invalid QR – not a visitor pass.' })
//   enqueue(async () => {
//     const { wb, ws } = await openBook()
//     let dup = false
//     ws.eachRow((row, i) => { if (i > 1 && row.getCell(2).value === v.registrationId) dup = true })
//     if (dup) return res.status(409).json({ ok: false, duplicate: true, error: 'Already scanned earlier.', visitor: v })
//     const scannedAt = fmtIST(Date.now())
//     ws.addRow([ws.rowCount, v.registrationId, v.name, v.visitorCount, v.personToMeet, v.visitDateTime, scannedAt])
//     const tmp = FILE + '.tmp'
//     await wb.xlsx.writeFile(tmp)
//     fs.renameSync(tmp, FILE)
//     res.json({ ok: true, visitor: { ...v, scannedAt } })
//   }).catch((e) => {
//     console.error(e)
//     if (!res.headersSent) res.status(500).json({ ok: false, error: 'Could not write Excel. Is visitors.xlsx open in Excel? Close it and retry.' })
//   })
// })

// app.get('/api/visitors', (_req, res) => {
//   enqueue(async () => {
//     const { ws } = await openBook()
//     const rows = []
//     ws.eachRow((r, i) => { if (i > 1) rows.push(COLUMNS.map((_, c) => r.getCell(c + 1).value)) })
//     res.json({ columns: COLUMNS.map((c) => c.header), rows: rows.reverse() })
//   }).catch(() => res.status(500).json({ error: 'read failed' }))
// })

// app.get('/download', (_req, res) => {
//   enqueue(async () => {
//     const { wb } = await openBook()
//     if (!fs.existsSync(FILE)) await wb.xlsx.writeFile(FILE)
//     res.download(FILE, 'visitors.xlsx')
//   })
// })

// async function start() {
//   let server
//   let proto = 'http'
//   if (process.env.HTTPS) {
//     const selfsigned = (await import('selfsigned')).default
//     const pems = selfsigned.generate([{ name: 'commonName', value: 'localhost' }], { days: 365, keySize: 2048 })
//     server = https.createServer({ key: pems.private, cert: pems.cert }, app)
//     proto = 'https'
//   } else server = http.createServer(app)
//   server.listen(PORT, '0.0.0.0', () => {
//     console.log(`\nVisitor QR Scanner running`)
//     console.log(`  This computer : ${proto}://localhost:${PORT}`)
//     for (const list of Object.values(os.networkInterfaces()))
//       for (const n of list || []) if (n.family === 'IPv4' && !n.internal) console.log(`  On your phone : ${proto}://${n.address}:${PORT}`)
//     console.log(`  Excel file    : ${FILE}\n`)
//   })
// }
// start()
import express from 'express'
import ExcelJS from 'exceljs'
import fs from 'fs'
import path from 'path'
import http from 'http'
import https from 'https'
import os from 'os'
import { fileURLToPath } from 'url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const PORT = process.env.PORT || 3000
const ON_VERCEL = !!process.env.VERCEL
// New private Blob stores authenticate via OIDC + BLOB_STORE_ID (no token); older ones use BLOB_READ_WRITE_TOKEN
const USE_BLOB = !!(process.env.BLOB_READ_WRITE_TOKEN || process.env.BLOB_STORE_ID)
const FILE = path.join(__dirname, 'data', 'visitors.xlsx')
const BLOB_PATH = 'visitors.xlsx'
const SHEET = 'Visitors'
const COLUMNS = [
  { header: 'Sr No', key: 'sr', width: 8 },
  { header: 'Registration ID', key: 'registrationId', width: 24 },
  { header: 'Name', key: 'name', width: 26 },
  { header: 'No. of Visitors', key: 'visitorCount', width: 16 },
  { header: 'Person to Meet', key: 'personToMeet', width: 24 },
  { header: 'Visit Date & Time', key: 'visitDateTime', width: 22 },
  { header: 'Scanned At', key: 'scannedAt', width: 22 },
]

/* ---------- Storage: Vercel Blob (online) or local file (your computer) ---------- */
// load() -> { buf: Buffer|null, etag }      save(buf, etag) -> throws { conflict:true } if file changed meanwhile
const storage = USE_BLOB ? await blobStorage() : ON_VERCEL ? missingBlobStorage() : fileStorage()

// Vercel's filesystem is read-only, so without Blob the Excel file can never be saved. Fail with a clear message.
function missingBlobStorage() {
  const fail = async () => { throw new Error('Vercel Blob is not connected (no BLOB_STORE_ID / BLOB_READ_WRITE_TOKEN). In Vercel: Storage -> create a PRIVATE Blob store -> connect it to this project -> Redeploy.') }
  return { name: 'NOT CONFIGURED (Vercel Blob missing)', load: fail, save: fail }
}

function fileStorage() {
  return {
    name: FILE,
    async load() { return { buf: fs.existsSync(FILE) ? fs.readFileSync(FILE) : null, etag: null } },
    async save(buf) {
      fs.mkdirSync(path.dirname(FILE), { recursive: true })
      fs.writeFileSync(FILE + '.tmp', buf)
      fs.renameSync(FILE + '.tmp', FILE)
    },
  }
}

async function blobStorage() {
  const { get, put, BlobPreconditionFailedError } = await import('@vercel/blob')
  return {
    name: 'Vercel Blob (private) → ' + BLOB_PATH,
    async load() {
      let r
      try { r = await get(BLOB_PATH, { access: 'private', useCache: false }) }
      catch (e) { if (/not.?found|404/i.test(String(e?.name) + ' ' + String(e?.message))) return { buf: null, etag: null }; throw e }
      if (!r || r.statusCode === 404) return { buf: null, etag: null }
      return { buf: Buffer.from(await new Response(r.stream).arrayBuffer()), etag: r.blob.etag }
    },
    async save(buf, etag) {
      try {
        await put(BLOB_PATH, buf, {
          access: 'private', addRandomSuffix: false, allowOverwrite: !!etag, ...(etag ? { ifMatch: etag } : {}),
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', cacheControlMaxAge: 60,
        })
      } catch (e) {
        if (e instanceof BlobPreconditionFailedError || /already exists/i.test(String(e?.message))) throw Object.assign(new Error('conflict'), { conflict: true })
        throw e
      }
    },
  }
}

if (process.env.STORAGE === 'memory') { // dev/test only
  let mem = null, v = 0
  Object.assign(storage, {
    async load() { return { buf: mem, etag: String(v) } },
    async save(buf, etag) { if (etag !== String(v)) throw Object.assign(new Error('conflict'), { conflict: true }); mem = buf; v++ },
  })
}

/* ---------- Excel helpers ---------- */
const fmtIST = (d) => new Date(d).toLocaleString('en-IN', { timeZone: 'Asia/Kolkata', dateStyle: 'medium', timeStyle: 'medium' })

async function openBook() {
  const { buf, etag } = await storage.load()
  const wb = new ExcelJS.Workbook()
  if (buf) {
    await wb.xlsx.load(buf)
    return { wb, ws: wb.getWorksheet(SHEET), etag }
  }
  const ws = wb.addWorksheet(SHEET, { views: [{ state: 'frozen', ySplit: 1 }] })
  ws.columns = COLUMNS
  const h = ws.getRow(1)
  h.font = { bold: true, color: { argb: 'FFFFFFFF' } }
  h.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F7A4D' } }
  h.alignment = { vertical: 'middle' }
  return { wb, ws, etag }
}

// Accepts: new compact QR {id,n,c,m,t}, old full QR {registrationId,name,...}, JSON surrounded by noise,
// URL-encoded JSON, or just a plain registration ID (VIS-2026-XXXXXXX) / URL containing it.
function parsePayload(raw) {
  const s = (v) => (v == null ? '' : String(v).slice(0, 200))
  let text = typeof raw === 'string' ? raw : raw && typeof raw === 'object' ? JSON.stringify(raw) : ''
  text = text.replace(/^\uFEFF/, '').trim()
  if (!text) return null

  const tryJson = (t) => { try { const o = JSON.parse(t); return o && typeof o === 'object' ? o : null } catch { return null } }
  let d = tryJson(text)
  if (!d) { try { d = tryJson(decodeURIComponent(text)) } catch {} }
  if (!d) { const m = text.match(/\{[\s\S]*\}/); if (m) d = tryJson(m[0]) }

  if (d) {
    const id = s(d.id ?? d.registrationId ?? d.regId ?? d.visitorId).trim()
    if (!id) return null
    return {
      registrationId: id,
      name: s(d.n ?? d.name),
      visitorCount: Number(d.c ?? d.visitorCount ?? d.count) || 1,
      personToMeet: s(d.m ?? d.personToMeet ?? d.meet),
      visitDateTime: s(d.t ?? d.visitDateTime ?? d.when).replace('T', ' '),
    }
  }
  // Plain ID (or a URL / text containing it)
  const m = text.match(/VIS-\d{4}-[A-Z0-9]+/i)
  if (m) return { registrationId: m[0].toUpperCase(), name: '', visitorCount: 1, personToMeet: '', visitDateTime: '' }
  return null
}

// Serialize writes inside one server instance
let queue = Promise.resolve()
const enqueue = (fn) => (queue = queue.then(fn, fn))

// Add one visitor; retries if another request changed the file at the same moment
async function addVisitor(v) {
  for (let i = 0; i < 6; i++) {
    const { wb, ws, etag } = await openBook()
    let dup = false
    ws.eachRow((row, n) => { if (n > 1 && row.getCell(2).value === v.registrationId) dup = true })
    if (dup) return { duplicate: true }
    const scannedAt = fmtIST(Date.now())
    ws.addRow([ws.rowCount, v.registrationId, v.name, v.visitorCount, v.personToMeet, v.visitDateTime, scannedAt])
    try {
      await storage.save(Buffer.from(await wb.xlsx.writeBuffer()), etag)
      return { scannedAt }
    } catch (e) { if (!e.conflict) throw e }
  }
  throw new Error('Too many simultaneous scans, please scan again.')
}

/* ---------- Web app ---------- */
const app = express()
app.use(express.json({ limit: '100kb' }))
app.use(express.static(path.join(__dirname, 'public')))

app.post('/api/scan', (req, res) => {
  const v = parsePayload(req.body?.qr)
  if (!v) return res.status(400).json({ ok: false, error: 'Invalid QR – not a visitor pass.', received: String(req.body?.qr ?? '').slice(0, 120) })
  enqueue(async () => {
    const r = await addVisitor(v)
    if (r.duplicate) return res.status(409).json({ ok: false, duplicate: true, error: 'Already scanned earlier.', visitor: v })
    res.json({ ok: true, visitor: { ...v, scannedAt: r.scannedAt } })
  }).catch((e) => {
    console.error(e)
    if (!res.headersSent) res.status(500).json({ ok: false, detail: String(e?.message || e).slice(0, 300), error: ON_VERCEL || USE_BLOB ? 'Could not save entry. Check Blob storage is connected.' : 'Could not write Excel. Is visitors.xlsx open in Excel? Close it and retry.' })
  })
})

app.get('/api/visitors', (_req, res) => {
  enqueue(async () => {
    const { ws } = await openBook()
    const rows = []
    ws.eachRow((r, i) => { if (i > 1) rows.push(COLUMNS.map((_, c) => r.getCell(c + 1).value)) })
    res.set('Cache-Control', 'no-store').json({ columns: COLUMNS.map((c) => c.header), rows: rows.reverse() })
  }).catch((e) => { console.error(e); res.status(500).json({ error: 'read failed' }) })
})

app.get('/download', (_req, res) => {
  enqueue(async () => {
    const { wb } = await openBook()
    const buf = Buffer.from(await wb.xlsx.writeBuffer())
    res.set({ 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': 'attachment; filename="visitors.xlsx"', 'Cache-Control': 'no-store' }).send(buf)
  }).catch((e) => { console.error(e); res.status(500).send('download failed') })
})

export default app

/* ---------- Local run (npm start). Vercel imports the app above instead. ---------- */
if (!ON_VERCEL) {
  let server, proto = 'http'
  if (process.env.HTTPS) {
    const selfsigned = (await import('selfsigned')).default
    const pems = selfsigned.generate([{ name: 'commonName', value: 'localhost' }], { days: 365, keySize: 2048 })
    server = https.createServer({ key: pems.private, cert: pems.cert }, app)
    proto = 'https'
  } else server = http.createServer(app)
  server.listen(PORT, '0.0.0.0', () => {
    console.log(`\nVisitor QR Scanner running`)
    console.log(`  This computer : ${proto}://localhost:${PORT}`)
    for (const list of Object.values(os.networkInterfaces()))
      for (const n of list || []) if (n.family === 'IPv4' && !n.internal) console.log(`  On your phone : ${proto}://${n.address}:${PORT}`)
    console.log(`  Excel file    : ${storage.name}\n`)
  })
}