import { Router } from 'express'
import { requireAuth, requireAdmin } from '../middleware/auth.js'
import { supabase } from '../lib/supabase.js'

const router = Router()

router.use(requireAuth)
router.use(requireAdmin)

// GET /api/admin/users
router.get('/users', async (req, res) => {
  try {
    const page = Number(req.query.page || 1)
    const limit = Number(req.query.limit || 10)
    const search = req.query.search || ''

    const from = (page - 1) * limit
    const to = from + limit - 1

    let query = supabase
      .from('users')
      .select(
        'id, name, email, role, created_at',
        { count: 'exact' }
      )

    if (search) {
      query = query.or(
        `name.ilike.%${search}%,email.ilike.%${search}%`
      )
    }

    const { data, error, count } =
      await query.range(from, to)

    if (error) throw error

    res.json({
      users: data,
      total: count
    })

  } catch (err) {
    res.status(500).json({
      error: err.message
    })
  }
})

export default router
