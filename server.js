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
const FILE = path.join(__dirname, 'data', 'visitors.xlsx')
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

const app = express()
app.use(express.json({ limit: '100kb' }))
app.use(express.static(path.join(__dirname, 'public')))
app.use('/lib/html5-qrcode.min.js', express.static(path.join(__dirname, 'node_modules/html5-qrcode/html5-qrcode.min.js')))

// Serialize writes so two quick scans never corrupt the file
let queue = Promise.resolve()
const enqueue = (fn) => (queue = queue.then(fn, fn))

const fmtIST = (d) => new Date(d).toLocaleString('en-IN', { timeZone: 'Asia/Kolkata', dateStyle: 'medium', timeStyle: 'medium' })

async function openBook() {
  const wb = new ExcelJS.Workbook()
  if (fs.existsSync(FILE)) {
    await wb.xlsx.readFile(FILE)
    return { wb, ws: wb.getWorksheet(SHEET) }
  }
  const ws = wb.addWorksheet(SHEET, { views: [{ state: 'frozen', ySplit: 1 }] })
  ws.columns = COLUMNS
  const h = ws.getRow(1)
  h.font = { bold: true, color: { argb: 'FFFFFFFF' } }
  h.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F7A4D' } }
  h.alignment = { vertical: 'middle' }
  return { wb, ws }
}

function parsePayload(raw) {
  let d
  try { d = typeof raw === 'string' ? JSON.parse(raw) : raw } catch { return null }
  if (!d || typeof d !== 'object' || !d.registrationId) return null
  const s = (v) => (v == null ? '' : String(v).slice(0, 200))
  return {
    registrationId: s(d.registrationId).trim(),
    name: s(d.name),
    visitorCount: Number(d.visitorCount) || 1,
    personToMeet: s(d.personToMeet),
    visitDateTime: s(d.visitDateTime).replace('T', ' '),
  }
}

app.post('/api/scan', (req, res) => {
  const v = parsePayload(req.body?.qr)
  if (!v) return res.status(400).json({ ok: false, error: 'Invalid QR – not a visitor pass.' })
  enqueue(async () => {
    const { wb, ws } = await openBook()
    let dup = false
    ws.eachRow((row, i) => { if (i > 1 && row.getCell(2).value === v.registrationId) dup = true })
    if (dup) return res.status(409).json({ ok: false, duplicate: true, error: 'Already scanned earlier.', visitor: v })
    const scannedAt = fmtIST(Date.now())
    ws.addRow([ws.rowCount, v.registrationId, v.name, v.visitorCount, v.personToMeet, v.visitDateTime, scannedAt])
    const tmp = FILE + '.tmp'
    await wb.xlsx.writeFile(tmp)
    fs.renameSync(tmp, FILE)
    res.json({ ok: true, visitor: { ...v, scannedAt } })
  }).catch((e) => {
    console.error(e)
    if (!res.headersSent) res.status(500).json({ ok: false, error: 'Could not write Excel. Is visitors.xlsx open in Excel? Close it and retry.' })
  })
})

app.get('/api/visitors', (_req, res) => {
  enqueue(async () => {
    const { ws } = await openBook()
    const rows = []
    ws.eachRow((r, i) => { if (i > 1) rows.push(COLUMNS.map((_, c) => r.getCell(c + 1).value)) })
    res.json({ columns: COLUMNS.map((c) => c.header), rows: rows.reverse() })
  }).catch(() => res.status(500).json({ error: 'read failed' }))
})

app.get('/download', (_req, res) => {
  enqueue(async () => {
    const { wb } = await openBook()
    if (!fs.existsSync(FILE)) await wb.xlsx.writeFile(FILE)
    res.download(FILE, 'visitors.xlsx')
  })
})

async function start() {
  let server
  let proto = 'http'
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
    console.log(`  Excel file    : ${FILE}\n`)
  })
}
start()