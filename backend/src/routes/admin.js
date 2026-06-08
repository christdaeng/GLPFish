import { Router } from 'express'
import { requireAuth, requireAdmin } from '../middleware/auth.js'
import { supabase } from '../lib/supabase.js'

const router = Router()

router.use(requireAuth)
router.use(requireAdmin)

// GET /api/admin/users
router.get('/users', async (req, res) => {
  try {
    const page = Math.max(Number(req.query.page) || 1, 1)
    const limit = Math.min(Math.max(Number(req.query.limit) || 10, 1), 100)
    const search = String(req.query.search || '').trim()

    const from = (page - 1) * limit
    const to = from + limit - 1

    let query = supabase
      .from('users')
      .select('id, name, email, role, created_at', { count: 'exact' })
      .order('created_at', { ascending: false })

    if (search) {
      query = query.or(`name.ilike.%${search}%,email.ilike.%${search}%`)
    }

    const { data, error, count } = await query.range(from, to)

    if (error) throw error

    res.json({
      users: data || [],
      total: count || 0,
      page,
      limit,
      totalPages: Math.ceil((count || 0) / limit)
    })
  } catch (err) {
    res.status(500).json({
      error: err.message
    })
  }
})

// GET /api/admin/users/:id
router.get('/users/:id', async (req, res) => {
  try {
    const { id } = req.params

    const { data: user, error: userError } = await supabase
      .from('users')
      .select('id, name, email, role, created_at')
      .eq('id', id)
      .single()

    if (userError) {
      if (userError.code === 'PGRST116') {
        return res.status(404).json({
          error: 'User tidak ditemukan'
        })
      }

      throw userError
    }

    const { data: inspections, error: inspectionError } = await supabase
      .from('images')
      .select(`
        id,
        file_name,
        uploaded_at,
        prediction_results (
          grade,
          label_text,
          confidence_score,
          predicted_at
        )
      `)
      .eq('user_id', id)
      .order('uploaded_at', { ascending: false })
      .limit(10)

    if (inspectionError) throw inspectionError

    res.json({
      user: {
        ...user,
        recentInspections: inspections || []
      }
    })
  } catch (err) {
    res.status(500).json({
      error: err.message
    })
  }
})

export default router
