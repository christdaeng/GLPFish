import {
  getSummaryData,
  getChartData,
  getTopUsersData,
  getRecentBatchesData,
  getActivityData,
  getUsersData,
  getUserDetailData,
  getRecapFile
} from '../services/adminService.js'

function getErrorStatus(error) {
  return error.statusCode || error.status || 500
}

function sendError(res, error) {
  console.error('[ADMIN ERROR]', error)

  res.status(getErrorStatus(error)).json({
    error: error.message || 'Terjadi kesalahan pada server'
  })
}

// GET /api/admin/summary?dari=&sampai=
export async function getAdminSummary(req, res) {
  try {
    const summary = await getSummaryData(req.query)

    res.json({
      summary
    })
  } catch (error) {
    sendError(res, error)
  }
}

// GET /api/admin/chart?dari=&sampai=
export async function getAdminChart(req, res) {
  try {
    const chart = await getChartData(req.query)

    res.json({
      chart
    })
  } catch (error) {
    sendError(res, error)
  }
}

// GET /api/admin/top-users?limit=5
export async function getTopUsers(req, res) {
  try {
    const users = await getTopUsersData(req.query)

    res.json({
      users
    })
  } catch (error) {
    sendError(res, error)
  }
}

// GET /api/admin/recent-batches?limit=8
export async function getRecentBatches(req, res) {
  try {
    const batches = await getRecentBatchesData(req.query)

    res.json({
      batches
    })
  } catch (error) {
    sendError(res, error)
  }
}

// GET /api/admin/activity?limit=10
export async function getAdminActivity(req, res) {
  try {
    const activities = await getActivityData(req.query)

    res.json({
      activities
    })
  } catch (error) {
    sendError(res, error)
  }
}

// GET /api/admin/recap?dari=&sampai=&jenis=&status=
export async function exportAdminRecap(req, res) {
  try {
    const { buffer, fileName } = await getRecapFile(req.query)

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${fileName}"`
    )

    res.send(buffer)
  } catch (error) {
    sendError(res, error)
  }
}

// GET /api/admin/users?search=&status=&page=&limit=
export async function getAdminUsers(req, res) {
  try {
    const result = await getUsersData(req.query)

    res.json(result)
  } catch (error) {
    sendError(res, error)
  }
}

// GET /api/admin/users/:id
export async function getAdminUserDetail(req, res) {
  try {
    const user = await getUserDetailData(req.params.id)

    res.json({
      user
    })
  } catch (error) {
    sendError(res, error)
  }
}
