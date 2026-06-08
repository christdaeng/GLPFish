import express from 'express'
import cors from 'cors'
import helmet from 'helmet'
import morgan from 'morgan'
import dotenv from 'dotenv'
import rateLimit from 'express-rate-limit'
import authRoutes       from './routes/auth.js'
import userRoutes       from './routes/user.js'
import uploadRoutes     from './routes/upload.js'
import inspectionRoutes from './routes/inspections.js'
import adminRoutes      from './routes/admin.js'

dotenv.config({ path: '.env.local' })

const app = express()

// ── Security & Logging ────────────────────────────────────
app.use(helmet())
app.use(morgan('dev'))

// ── Rate limiting ─────────────────────────────────────────
app.use('/api/auth', rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 20,
  message: { error: 'Terlalu banyak percobaan, coba lagi nanti' },
}))

// ── CORS ──────────────────────────────────────────────────
app.use(cors({
  origin: 'http://localhost:5173',
  credentials: true,
}))

// ── Body Parser ───────────────────────────────────────────
// Skip express.json() untuk route upload — multer yang handle body parsing-nya
app.use((req, res, next) => {
  if (req.originalUrl.startsWith('/api/upload')) return next()
  express.json()(req, res, next)
})

// ── Routes ────────────────────────────────────────────────
app.use('/api/auth',        authRoutes)
app.use('/api/users',        userRoutes)
app.use('/api/upload',      uploadRoutes)
app.use('/api/inspections', inspectionRoutes)
app.use('/api/admin', adminRoutes)

// ── Health check ──────────────────────────────────────────
app.get('/api/health', (_, res) => res.json({ status: 'ok' }))

// ── Global error handler ──────────────────────────────────
app.use((err, req, res, next) => {
  console.error('[ERROR]', err.message)
  res.status(err.status ?? 500).json({ error: err.message ?? 'Internal server error' })
})

// ── 404 handler ───────────────────────────────────────────
app.use((_, res) => res.status(404).json({ error: 'Route tidak ditemukan' }))

const PORT = process.env.PORT ?? 4000
app.listen(PORT, () => console.log(`✅ Backend running → http://localhost:${PORT}`))
