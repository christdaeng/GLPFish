import { Router } from 'express'
import { requireAuth, requireAdmin } from '../middleware/auth.js'
import {
  getAdminSummary,
  getAdminChart,
  getTopUsers,
  getRecentBatches,
  getAdminActivity,
  exportAdminRecap,
  getAdminUsers,
  getAdminUserDetail
} from '../controllers/adminController.js'

const router = Router()

router.use(requireAuth)
router.use(requireAdmin)

// GET /api/admin/summary?dari=&sampai=
router.get('/summary', getAdminSummary)

// GET /api/admin/chart?dari=&sampai=
router.get('/chart', getAdminChart)

// GET /api/admin/top-users?limit=5
router.get('/top-users', getTopUsers)

// GET /api/admin/recent-batches?limit=8
router.get('/recent-batches', getRecentBatches)

// GET /api/admin/activity?limit=10
router.get('/activity', getAdminActivity)

// GET /api/admin/recap?dari=&sampai=&jenis=&status=
router.get('/recap', exportAdminRecap)

// GET /api/admin/users?search=&status=&page=&limit=
router.get('/users', getAdminUsers)

// GET /api/admin/users/:id
router.get('/users/:id', getAdminUserDetail)

export default router
