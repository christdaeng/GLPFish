import * as XLSX from 'xlsx'
import { supabase } from '../lib/supabase.js'

function todayIso() {
  return new Date().toISOString().slice(0, 10)
}

function daysAgoIso(days) {
  return new Date(Date.now() - days * 24 * 60 * 60 * 1000)
    .toISOString()
    .slice(0, 10)
}

function startOfDay(date) {
  return `${date}T00:00:00.000Z`
}

function endOfDay(date) {
  return `${date}T23:59:59.999Z`
}

function getDateRange(query = {}) {
  const dari = String(query.dari || '').trim() || daysAgoIso(30)
  const sampai = String(query.sampai || '').trim() || todayIso()

  return { dari, sampai }
}

function formatDate(value) {
  if (!value) return '-'

  return new Date(value).toLocaleDateString('id-ID', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  })
}

function formatDateTime(value) {
  if (!value) return '-'

  return new Date(value).toLocaleString('id-ID')
}

function normalizeGrade(value) {
  if (!value) return ''

  return String(value)
    .replace(/^grade\s*/i, '')
    .trim()
    .toUpperCase()
}

function calculatePercent(part, total) {
  if (!total) return 0
  return Math.round((part / total) * 100)
}

function unique(values) {
  return [...new Set(values.filter(Boolean))]
}

function latestDate(...values) {
  const timestamps = values
    .filter(Boolean)
    .map((value) => new Date(value).getTime())
    .filter((value) => Number.isFinite(value))

  if (timestamps.length === 0) return null

  return new Date(Math.max(...timestamps)).toISOString()
}

function average(values) {
  const numbers = values.filter((value) => {
    return typeof value === 'number' && !Number.isNaN(value)
  })

  if (numbers.length === 0) return null

  return numbers.reduce((sum, value) => sum + value, 0) / numbers.length
}

function getUserStatus(lastActiveRaw) {
  if (!lastActiveRaw) return 'Tidak Aktif'

  const inactiveDays = Math.floor(
    (Date.now() - new Date(lastActiveRaw).getTime()) /
      (24 * 60 * 60 * 1000)
  )

  return inactiveDays <= 30 ? 'Aktif' : 'Tidak Aktif'
}

function matchUserStatus(userStatus, statusFilter) {
  if (!statusFilter || statusFilter === 'semua') return true

  const filter = String(statusFilter).trim().toLowerCase()

  if (filter === 'aktif') return userStatus === 'Aktif'

  if (
    filter === 'tidak aktif' ||
    filter === 'nonaktif' ||
    filter === 'inactive'
  ) {
    return userStatus === 'Tidak Aktif'
  }

  return true
}

function normalizeBatchStatus(status) {
  const value = String(status || '').trim().toLowerCase()

  const labels = {
    pending: 'Menunggu',
    processing: 'Diproses',
    done: 'Selesai',
    failed: 'Gagal'
  }

  return labels[value] || status || '-'
}

function getInspectionStatus(fish, prediction) {
  const grade = normalizeGrade(prediction?.grade)

  if (prediction?.exportable === false || grade === 'C') {
    return 'Reject'
  }

  if (prediction) return 'Selesai'

  if (fish?.status === 'failed') return 'Gagal'
  if (fish?.status === 'processing') return 'Diproses'

  return 'Menunggu'
}

function matchStatusValue(rowStatus, filter) {
  if (!filter || filter === 'semua') return true

  const left = String(rowStatus || '').trim().toLowerCase()
  const right = String(filter || '').trim().toLowerCase()

  if (!right || right === 'semua') return true
  if (left === right) return true

  const aliases = {
    selesai: ['done', 'selesai'],
    menunggu: ['pending', 'menunggu'],
    diproses: ['processing', 'diproses'],
    gagal: ['failed', 'gagal'],
    reject: ['reject', 'ditolak']
  }

  return Object.values(aliases).some((items) => {
    return items.includes(left) && items.includes(right)
  })
}

function createWorkbookBuffer(sheets) {
  const workbook = XLSX.utils.book_new()

  sheets.forEach(({ name, rows }) => {
    const worksheet = XLSX.utils.json_to_sheet(rows)
    XLSX.utils.book_append_sheet(workbook, worksheet, name)
  })

  return XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'buffer'
  })
}

async function fetchUsersByIds(userIds) {
  const ids = unique(userIds)
  if (ids.length === 0) return []

  const { data, error } = await supabase
    .from('users')
    .select('id, name, email, role, created_at')
    .in('id', ids)

  if (error) throw error

  return data || []
}

async function fetchBatchesByIds(batchIds) {
  const ids = unique(batchIds)
  if (ids.length === 0) return []

  const { data, error } = await supabase
    .from('batches')
    .select(`
      id,
      user_id,
      fish_count,
      fish_category,
      status,
      confidence_score_avg,
      created_at,
      submitted_at,
      lokasi,
      berat_total,
      catatan,
      preprocessed_status
    `)
    .in('id', ids)

  if (error) throw error

  return data || []
}

async function fetchFishesByBatchIds(batchIds) {
  const ids = unique(batchIds)
  if (ids.length === 0) return []

  const { data, error } = await supabase
    .from('fishes')
    .select(`
      id,
      batch_id,
      fish_index,
      eye_image_id,
      gill_image_id,
      status,
      created_at
    `)
    .in('batch_id', ids)

  if (error) throw error

  return data || []
}

async function fetchFishesByIds(fishIds) {
  const ids = unique(fishIds)
  if (ids.length === 0) return []

  const { data, error } = await supabase
    .from('fishes')
    .select(`
      id,
      batch_id,
      fish_index,
      eye_image_id,
      gill_image_id,
      status,
      created_at
    `)
    .in('id', ids)

  if (error) throw error

  return data || []
}

async function fetchPredictionsByFishIds(fishIds) {
  const ids = unique(fishIds)
  if (ids.length === 0) return []

  const { data, error } = await supabase
    .from('prediction_results')
    .select(`
      id,
      fish_id,
      confidence_score,
      grade,
      label_text,
      exportable,
      eyes_status,
      eyes_confidence_score,
      gill_status,
      gill_confidence_score,
      waktu_proses_ms,
      predicted_at
    `)
    .in('fish_id', ids)

  if (error) throw error

  return data || []
}

async function fetchImagesByIds(imageIds) {
  const ids = unique(imageIds)
  if (ids.length === 0) return []

  const { data, error } = await supabase
    .from('images')
    .select(`
      id,
      user_id,
      storage_path,
      file_name,
      mime_type,
      file_size,
      uploaded_at
    `)
    .in('id', ids)

  if (error) throw error

  return data || []
}

function buildBatchStats(batches, fishes, predictions) {
  const fishesByBatch = new Map()
  const predictionByFish = new Map()

  fishes.forEach((fish) => {
    if (!fishesByBatch.has(fish.batch_id)) {
      fishesByBatch.set(fish.batch_id, [])
    }

    fishesByBatch.get(fish.batch_id).push(fish)
  })

  predictions.forEach((prediction) => {
    predictionByFish.set(prediction.fish_id, prediction)
  })

  return batches.map((batch) => {
    const batchFishes = fishesByBatch.get(batch.id) || []

    const batchPredictions = batchFishes
      .map((fish) => predictionByFish.get(fish.id))
      .filter(Boolean)

    const totalInspeksi = batchFishes.length
    const totalPrediksi = batchPredictions.length

    const totalGradeA = batchPredictions.filter((prediction) => {
      return normalizeGrade(prediction.grade) === 'A'
    }).length

    const totalReject = batchPredictions.filter((prediction) => {
      const grade = normalizeGrade(prediction.grade)
      return grade === 'C' || prediction.exportable === false
    }).length

    const confidenceScoreAvg =
      typeof batch.confidence_score_avg === 'number'
        ? batch.confidence_score_avg
        : average(batchPredictions.map((prediction) => prediction.confidence_score))

    return {
      ...batch,
      fishes: batchFishes,
      predictions: batchPredictions,
      totalInspeksi,
      totalPrediksi,
      totalGradeA,
      totalReject,
      gradeAPercent: calculatePercent(totalGradeA, totalPrediksi || totalInspeksi),
      rejectPercent: calculatePercent(totalReject, totalPrediksi || totalInspeksi),
      confidenceScoreAvg
    }
  })
}

function toBatchResponse(batch, user) {
  return {
    id: batch.id,
    userId: batch.user_id,
    userName: user?.name || user?.email || 'Tidak diketahui',
    userEmail: user?.email || '-',
    fishCount: batch.fish_count,
    fishCategory: batch.fish_category || '-',
    status: batch.status,
    statusLabel: normalizeBatchStatus(batch.status),
    confidenceScoreAvg: batch.confidenceScoreAvg ?? batch.confidence_score_avg ?? null,
    createdAt: batch.created_at,
    submittedAt: batch.submitted_at,
    tanggal: formatDate(batch.created_at),
    lokasi: batch.lokasi,
    beratTotal: batch.berat_total,
    catatan: batch.catatan || '-',
    preprocessedStatus: batch.preprocessed_status,
    totalInspeksi: batch.totalInspeksi || 0,
    totalPrediksi: batch.totalPrediksi || 0,
    gradeA: batch.totalGradeA || 0,
    reject: batch.totalReject || 0,
    gradeAPercent: batch.gradeAPercent || 0,
    rejectPercent: batch.rejectPercent || 0
  }
}

export async function getSummaryData(query = {}) {
  const { dari, sampai } = getDateRange(query)
  const awalBulanIni = `${todayIso().slice(0, 7)}-01`

  const [usersResult, newUsersResult, batchesResult] = await Promise.all([
    supabase
      .from('users')
      .select('id', { count: 'exact', head: true }),

    supabase
      .from('users')
      .select('id', { count: 'exact', head: true })
      .gte('created_at', startOfDay(awalBulanIni)),

    supabase
      .from('batches')
      .select(`
        id,
        user_id,
        fish_count,
        fish_category,
        status,
        confidence_score_avg,
        created_at,
        submitted_at,
        lokasi,
        berat_total,
        catatan,
        preprocessed_status
      `)
      .gte('created_at', startOfDay(dari))
      .lte('created_at', endOfDay(sampai))
  ])

  if (usersResult.error) throw usersResult.error
  if (newUsersResult.error) throw newUsersResult.error
  if (batchesResult.error) throw batchesResult.error

  const batches = batchesResult.data || []
  const fishes = await fetchFishesByBatchIds(batches.map((batch) => batch.id))
  const predictions = await fetchPredictionsByFishIds(fishes.map((fish) => fish.id))

  const totalBatch = batches.length
  const totalInspeksi = fishes.length
  const totalPrediksi = predictions.length

  const totalGradeA = predictions.filter((prediction) => {
    return normalizeGrade(prediction.grade) === 'A'
  }).length

  const totalReject = predictions.filter((prediction) => {
    const grade = normalizeGrade(prediction.grade)
    return grade === 'C' || prediction.exportable === false
  }).length

  return {
    totalPengguna: usersResult.count || 0,
    totalBatch,
    totalInspeksi,
    totalPrediksi,
    gradeAPercent: calculatePercent(totalGradeA, totalPrediksi || totalInspeksi),
    rejectPercent: calculatePercent(totalReject, totalPrediksi || totalInspeksi),
    newUsersThisMonth: newUsersResult.count || 0
  }
}

export async function getChartData(query = {}) {
  const { dari, sampai } = getDateRange(query)

  const { data: batchesData, error: batchesError } = await supabase
    .from('batches')
    .select(`
      id,
      user_id,
      fish_count,
      fish_category,
      status,
      confidence_score_avg,
      created_at,
      submitted_at,
      lokasi,
      berat_total,
      catatan,
      preprocessed_status
    `)
    .gte('created_at', startOfDay(dari))
    .lte('created_at', endOfDay(sampai))
    .order('created_at', { ascending: true })

  if (batchesError) throw batchesError

  const batches = batchesData || []
  const fishes = await fetchFishesByBatchIds(batches.map((batch) => batch.id))
  const predictions = await fetchPredictionsByFishIds(fishes.map((fish) => fish.id))

  const fishById = new Map(fishes.map((fish) => [fish.id, fish]))
  const batchById = new Map(batches.map((batch) => [batch.id, batch]))
  const grouped = {}

  batches.forEach((batch) => {
    const tanggal = new Date(batch.created_at).toISOString().slice(0, 10)

    if (!grouped[tanggal]) {
      grouped[tanggal] = {
        tanggal,
        totalBatch: 0,
        totalInspeksi: 0,
        totalPrediksi: 0,
        gradeA: 0,
        reject: 0
      }
    }

    grouped[tanggal].totalBatch += 1
  })

  fishes.forEach((fish) => {
    const batch = batchById.get(fish.batch_id)
    if (!batch) return

    const tanggal = new Date(batch.created_at).toISOString().slice(0, 10)
    if (!grouped[tanggal]) return

    grouped[tanggal].totalInspeksi += 1
  })

  predictions.forEach((prediction) => {
    const fish = fishById.get(prediction.fish_id)
    const batch = fish ? batchById.get(fish.batch_id) : null

    if (!batch) return

    const tanggal = new Date(batch.created_at).toISOString().slice(0, 10)
    const grade = normalizeGrade(prediction.grade)

    grouped[tanggal].totalPrediksi += 1

    if (grade === 'A') grouped[tanggal].gradeA += 1
    if (grade === 'C' || prediction.exportable === false) grouped[tanggal].reject += 1
  })

  return Object.values(grouped).map((item) => ({
    tanggal: item.tanggal,
    totalBatch: item.totalBatch,
    totalInspeksi: item.totalInspeksi,
    totalPrediksi: item.totalPrediksi,
    gradeA: item.gradeA,
    reject: item.reject,
    gradeAPercent: calculatePercent(item.gradeA, item.totalPrediksi || item.totalInspeksi),
    rejectPercent: calculatePercent(item.reject, item.totalPrediksi || item.totalInspeksi)
  }))
}

export async function getTopUsersData(query = {}) {
  const limit = Math.min(Math.max(Number(query.limit) || 5, 1), 20)

  const [usersResult, batchesResult] = await Promise.all([
    supabase
      .from('users')
      .select('id, name, email, role, created_at'),

    supabase
      .from('batches')
      .select(`
        id,
        user_id,
        fish_count,
        fish_category,
        status,
        confidence_score_avg,
        created_at,
        submitted_at,
        lokasi,
        berat_total,
        catatan,
        preprocessed_status
      `)
      .order('created_at', { ascending: false })
  ])

  if (usersResult.error) throw usersResult.error
  if (batchesResult.error) throw batchesResult.error

  const users = usersResult.data || []
  const batches = batchesResult.data || []
  const fishes = await fetchFishesByBatchIds(batches.map((batch) => batch.id))
  const predictions = await fetchPredictionsByFishIds(fishes.map((fish) => fish.id))

  const batchById = new Map(batches.map((batch) => [batch.id, batch]))
  const predictionByFish = new Map(predictions.map((prediction) => [prediction.fish_id, prediction]))
  const statsByUser = new Map()

  users.forEach((user) => {
    statsByUser.set(user.id, {
      totalBatch: 0,
      totalInspeksi: 0,
      totalPrediksi: 0,
      gradeA: 0,
      reject: 0,
      lastActiveRaw: user.created_at
    })
  })

  batches.forEach((batch) => {
    if (!statsByUser.has(batch.user_id)) return

    const stats = statsByUser.get(batch.user_id)

    stats.totalBatch += 1
    stats.lastActiveRaw = latestDate(stats.lastActiveRaw, batch.submitted_at, batch.created_at)
  })

  fishes.forEach((fish) => {
    const batch = batchById.get(fish.batch_id)
    if (!batch || !statsByUser.has(batch.user_id)) return

    const stats = statsByUser.get(batch.user_id)
    const prediction = predictionByFish.get(fish.id)
    const grade = normalizeGrade(prediction?.grade)

    stats.totalInspeksi += 1

    if (prediction) {
      stats.totalPrediksi += 1

      if (grade === 'A') stats.gradeA += 1
      if (grade === 'C' || prediction.exportable === false) stats.reject += 1

      stats.lastActiveRaw = latestDate(stats.lastActiveRaw, prediction.predicted_at)
    }
  })

  return users
    .map((user) => {
      const stats = statsByUser.get(user.id) || {
        totalBatch: 0,
        totalInspeksi: 0,
        totalPrediksi: 0,
        gradeA: 0,
        reject: 0,
        lastActiveRaw: user.created_at
      }

      return {
        id: user.id,
        name: user.name || '-',
        email: user.email || '-',
        role: user.role || '-',
        totalBatch: stats.totalBatch,
        totalInspeksi: stats.totalInspeksi,
        totalPrediksi: stats.totalPrediksi,
        gradeAPercent: calculatePercent(stats.gradeA, stats.totalPrediksi || stats.totalInspeksi),
        rejectPercent: calculatePercent(stats.reject, stats.totalPrediksi || stats.totalInspeksi),
        joined: formatDate(user.created_at),
        lastActive: formatDate(stats.lastActiveRaw),
        lastActiveRaw: stats.lastActiveRaw,
        status: getUserStatus(stats.lastActiveRaw)
      }
    })
    .sort((a, b) => b.totalInspeksi - a.totalInspeksi)
    .slice(0, limit)
}

export async function getRecentBatchesData(query = {}) {
  const limit = Math.min(Math.max(Number(query.limit) || 8, 1), 50)

  const { data: batchesData, error: batchesError } = await supabase
    .from('batches')
    .select(`
      id,
      user_id,
      fish_count,
      fish_category,
      status,
      confidence_score_avg,
      created_at,
      submitted_at,
      lokasi,
      berat_total,
      catatan,
      preprocessed_status
    `)
    .order('created_at', { ascending: false })
    .limit(limit)

  if (batchesError) throw batchesError

  const batches = batchesData || []
  const users = await fetchUsersByIds(batches.map((batch) => batch.user_id))
  const fishes = await fetchFishesByBatchIds(batches.map((batch) => batch.id))
  const predictions = await fetchPredictionsByFishIds(fishes.map((fish) => fish.id))

  const userMap = new Map(users.map((user) => [user.id, user]))
  const batchesWithStats = buildBatchStats(batches, fishes, predictions)

  return batchesWithStats.map((batch) => {
    return toBatchResponse(batch, userMap.get(batch.user_id))
  })
}
export async function getActivityData(query = {}) {
  const limit = Math.min(Math.max(Number(query.limit) || 10, 1), 50)

  const [usersResult, batchesResult, imagesResult, predictionsResult] = await Promise.all([
    supabase
      .from('users')
      .select('id, name, email, created_at')
      .order('created_at', { ascending: false })
      .limit(limit),

    supabase
      .from('batches')
      .select('id, user_id, fish_count, fish_category, status, created_at, submitted_at')
      .order('created_at', { ascending: false })
      .limit(limit),

    supabase
      .from('images')
      .select('id, user_id, file_name, uploaded_at')
      .order('uploaded_at', { ascending: false })
      .limit(limit),

    supabase
      .from('prediction_results')
      .select('id, fish_id, grade, label_text, confidence_score, exportable, predicted_at')
      .order('predicted_at', { ascending: false })
      .limit(limit)
  ])

  if (usersResult.error) throw usersResult.error
  if (batchesResult.error) throw batchesResult.error
  if (imagesResult.error) throw imagesResult.error
  if (predictionsResult.error) throw predictionsResult.error

  const users = usersResult.data || []
  const batches = batchesResult.data || []
  const images = imagesResult.data || []
  const predictions = predictionsResult.data || []

  const predictionFishes = await fetchFishesByIds(
    predictions.map((prediction) => prediction.fish_id)
  )

  const predictionBatches = await fetchBatchesByIds(
    predictionFishes.map((fish) => fish.batch_id)
  )

  const extraUsers = await fetchUsersByIds([
    ...batches.map((batch) => batch.user_id),
    ...images.map((image) => image.user_id),
    ...predictionBatches.map((batch) => batch.user_id)
  ])

  const userMap = new Map(extraUsers.map((user) => [user.id, user]))
  const fishMap = new Map(predictionFishes.map((fish) => [fish.id, fish]))
  const batchMap = new Map(predictionBatches.map((batch) => [batch.id, batch]))

  const activities = []

  users.forEach((user) => {
    activities.push({
      id: `user-${user.id}`,
      type: 'USER_CREATED',
      title: 'Pengguna baru terdaftar',
      description: `${user.name || user.email || 'Pengguna'} bergabung ke sistem`,
      userId: user.id,
      userName: user.name || user.email || 'Tidak diketahui',
      createdAt: user.created_at
    })
  })

  batches.forEach((batch) => {
    const user = userMap.get(batch.user_id)

    activities.push({
      id: `batch-${batch.id}`,
      type: 'BATCH_CREATED',
      title: 'Batch baru dibuat',
      description: `${user?.name || user?.email || 'Pengguna'} membuat batch ${batch.id}`,
      userId: batch.user_id,
      userName: user?.name || user?.email || 'Tidak diketahui',
      batchId: batch.id,
      status: batch.status,
      createdAt: batch.created_at
    })
  })

  images.forEach((image) => {
    const user = userMap.get(image.user_id)

    activities.push({
      id: `image-${image.id}`,
      type: 'IMAGE_UPLOADED',
      title: 'Gambar baru diunggah',
      description: `${user?.name || user?.email || 'Pengguna'} mengunggah ${image.file_name || 'gambar ikan'}`,
      userId: image.user_id,
      userName: user?.name || user?.email || 'Tidak diketahui',
      imageId: image.id,
      createdAt: image.uploaded_at
    })
  })

  predictions.forEach((prediction) => {
    const fish = fishMap.get(prediction.fish_id)
    const batch = fish ? batchMap.get(fish.batch_id) : null
    const user = batch ? userMap.get(batch.user_id) : null
    const grade = normalizeGrade(prediction.grade)

    activities.push({
      id: `prediction-${prediction.id}`,
      type: 'PREDICTION_CREATED',
      title: 'Hasil inspeksi tersedia',
      description: `Ikan ke-${fish?.fish_index ?? '-'} pada batch ${batch?.id || '-'} menghasilkan Grade ${grade || '-'}`,
      userId: batch?.user_id || null,
      userName: user?.name || user?.email || 'Tidak diketahui',
      batchId: batch?.id || null,
      fishId: prediction.fish_id,
      grade: grade || '-',
      label: prediction.label_text || '-',
      confidenceScore: prediction.confidence_score ?? null,
      createdAt: prediction.predicted_at
    })
  })

  return activities
    .filter((activity) => activity.createdAt)
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt))
    .slice(0, limit)
    .map((activity) => ({
      ...activity,
      tanggal: formatDate(activity.createdAt)
    }))
}

export async function getUsersData(query = {}) {
  const page = Math.max(Number(query.page) || 1, 1)
  const limit = Math.min(Math.max(Number(query.limit) || 10, 1), 100)
  const search = String(query.search || '').trim()
  const status = String(query.status || 'semua').trim().toLowerCase()

  let usersQuery = supabase
    .from('users')
    .select('id, name, email, role, created_at')
    .order('created_at', { ascending: false })

  if (search) {
    usersQuery = usersQuery.or(`name.ilike.%${search}%,email.ilike.%${search}%`)
  }

  const { data: usersData, error: usersError } = await usersQuery

  if (usersError) throw usersError

  const users = usersData || []
  const userIds = users.map((user) => user.id)

  const { data: batchesData, error: batchesError } =
    userIds.length > 0
      ? await supabase
          .from('batches')
          .select(`
            id,
            user_id,
            fish_count,
            fish_category,
            status,
            confidence_score_avg,
            created_at,
            submitted_at,
            lokasi,
            berat_total,
            catatan,
            preprocessed_status
          `)
          .in('user_id', userIds)
          .order('created_at', { ascending: false })
      : { data: [], error: null }

  if (batchesError) throw batchesError

  const batches = batchesData || []
  const fishes = await fetchFishesByBatchIds(batches.map((batch) => batch.id))
  const predictions = await fetchPredictionsByFishIds(fishes.map((fish) => fish.id))

  const batchById = new Map(batches.map((batch) => [batch.id, batch]))
  const predictionByFish = new Map(
    predictions.map((prediction) => [prediction.fish_id, prediction])
  )

  const statsByUser = new Map()

  users.forEach((user) => {
    statsByUser.set(user.id, {
      totalBatch: 0,
      totalInspeksi: 0,
      totalPrediksi: 0,
      gradeA: 0,
      reject: 0,
      lastActiveRaw: user.created_at
    })
  })

  batches.forEach((batch) => {
    if (!statsByUser.has(batch.user_id)) return

    const stats = statsByUser.get(batch.user_id)

    stats.totalBatch += 1
    stats.lastActiveRaw = latestDate(
      stats.lastActiveRaw,
      batch.submitted_at,
      batch.created_at
    )
  })

  fishes.forEach((fish) => {
    const batch = batchById.get(fish.batch_id)
    if (!batch || !statsByUser.has(batch.user_id)) return

    const stats = statsByUser.get(batch.user_id)
    const prediction = predictionByFish.get(fish.id)
    const grade = normalizeGrade(prediction?.grade)

    stats.totalInspeksi += 1

    if (prediction) {
      stats.totalPrediksi += 1

      if (grade === 'A') stats.gradeA += 1
      if (grade === 'C' || prediction.exportable === false) stats.reject += 1

      stats.lastActiveRaw = latestDate(stats.lastActiveRaw, prediction.predicted_at)
    }
  })

  const enrichedUsers = users
    .map((user) => {
      const stats = statsByUser.get(user.id) || {
        totalBatch: 0,
        totalInspeksi: 0,
        totalPrediksi: 0,
        gradeA: 0,
        reject: 0,
        lastActiveRaw: user.created_at
      }

      const userStatus = getUserStatus(stats.lastActiveRaw)

      return {
        id: user.id,
        name: user.name || '-',
        email: user.email || '-',
        role: user.role || '-',
        joined: formatDate(user.created_at),
        createdAt: user.created_at,
        totalBatch: stats.totalBatch,
        totalInspeksi: stats.totalInspeksi,
        totalPrediksi: stats.totalPrediksi,
        gradeAPercent: calculatePercent(
          stats.gradeA,
          stats.totalPrediksi || stats.totalInspeksi
        ),
        rejectPercent: calculatePercent(
          stats.reject,
          stats.totalPrediksi || stats.totalInspeksi
        ),
        lastActive: formatDate(stats.lastActiveRaw),
        lastActiveRaw: stats.lastActiveRaw,
        status: userStatus
      }
    })
    .filter((user) => matchUserStatus(user.status, status))

  const total = enrichedUsers.length
  const from = (page - 1) * limit
  const to = from + limit

  return {
    users: enrichedUsers.slice(from, to),
    total,
    page,
    limit,
    totalPages: Math.ceil(total / limit)
  }
}

export async function getUserDetailData(id) {
  const { data: user, error: userError } = await supabase
    .from('users')
    .select('id, name, email, role, created_at')
    .eq('id', id)
    .single()

  if (userError) {
    if (userError.code === 'PGRST116') {
      const error = new Error('User tidak ditemukan')
      error.statusCode = 404
      throw error
    }

    throw userError
  }

  const { data: batchesData, error: batchesError } = await supabase
    .from('batches')
    .select(`
      id,
      user_id,
      fish_count,
      fish_category,
      status,
      confidence_score_avg,
      created_at,
      submitted_at,
      lokasi,
      berat_total,
      catatan,
      preprocessed_status
    `)
    .eq('user_id', id)
    .order('created_at', { ascending: false })

  if (batchesError) throw batchesError

  const batches = batchesData || []
  const fishes = await fetchFishesByBatchIds(batches.map((batch) => batch.id))
  const predictions = await fetchPredictionsByFishIds(fishes.map((fish) => fish.id))
  const batchesWithStats = buildBatchStats(batches, fishes, predictions)

  const totalBatch = batchesWithStats.length
  const totalInspeksi = batchesWithStats.reduce(
    (sum, batch) => sum + batch.totalInspeksi,
    0
  )
  const totalPrediksi = batchesWithStats.reduce(
    (sum, batch) => sum + batch.totalPrediksi,
    0
  )
  const totalGradeA = batchesWithStats.reduce(
    (sum, batch) => sum + batch.totalGradeA,
    0
  )
  const totalReject = batchesWithStats.reduce(
    (sum, batch) => sum + batch.totalReject,
    0
  )

  const lastBatch = batchesWithStats[0] || null
  const lastActiveRaw = latestDate(
    lastBatch?.submitted_at,
    lastBatch?.created_at,
    user.created_at
  )

  const recentBatches = batchesWithStats.slice(0, 10).map((batch) => ({
    id: batch.id,
    batchId: batch.id,
    fishCount: batch.fish_count,
    fishCategory: batch.fish_category || '-',
    status: batch.status,
    statusLabel: normalizeBatchStatus(batch.status),
    confidenceScoreAvg: batch.confidenceScoreAvg ?? batch.confidence_score_avg ?? null,
    createdAt: batch.created_at,
    submittedAt: batch.submitted_at,
    tanggal: formatDate(batch.created_at),
    lokasi: batch.lokasi,
    beratTotal: batch.berat_total,
    preprocessedStatus: batch.preprocessed_status,
    totalInspeksi: batch.totalInspeksi,
    totalPrediksi: batch.totalPrediksi,
    gradeAPercent: batch.gradeAPercent,
    rejectPercent: batch.rejectPercent
  }))

  return {
    id: user.id,
    name: user.name || '-',
    email: user.email || '-',
    role: user.role || '-',
    joined: formatDate(user.created_at),
    createdAt: user.created_at,
    totalBatch,
    totalInspeksi,
    totalPrediksi,
    gradeAPercent: calculatePercent(totalGradeA, totalPrediksi || totalInspeksi),
    rejectPercent: calculatePercent(totalReject, totalPrediksi || totalInspeksi),
    lastActive: formatDate(lastActiveRaw),
    lastActiveRaw,
    status: getUserStatus(lastActiveRaw),
    recentBatches,
    recentInspections: recentBatches
  }
}

export async function getRecapFile(query = {}) {
  const { dari, sampai } = getDateRange(query)
  const jenis = String(query.jenis || 'inspeksi').trim().toLowerCase()
  const statusFilter = String(query.status || '').trim().toLowerCase()

  if (jenis === 'users' || jenis === 'pengguna') {
    const { data: users, error } = await supabase
      .from('users')
      .select('id, name, email, role, created_at')
      .gte('created_at', startOfDay(dari))
      .lte('created_at', endOfDay(sampai))
      .order('created_at', { ascending: false })

    if (error) throw error

    const rows = (users || []).map((user, index) => ({
      No: index + 1,
      'ID User': user.id,
      Nama: user.name || '-',
      Email: user.email || '-',
      Role: user.role || '-',
      'Tanggal Daftar': formatDateTime(user.created_at)
    }))

    return {
      buffer: createWorkbookBuffer([{ name: 'Rekap Pengguna', rows }]),
      fileName: `rekap-pengguna-${dari}-sampai-${sampai}.xlsx`
    }
  }

  const { data: batchesData, error: batchesError } = await supabase
    .from('batches')
    .select(`
      id,
      user_id,
      fish_count,
      fish_category,
      status,
      confidence_score_avg,
      created_at,
      submitted_at,
      lokasi,
      berat_total,
      catatan,
      preprocessed_status
    `)
    .gte('created_at', startOfDay(dari))
    .lte('created_at', endOfDay(sampai))
    .order('created_at', { ascending: false })

  if (batchesError) throw batchesError

  const batches = batchesData || []
  const users = await fetchUsersByIds(batches.map((batch) => batch.user_id))
  const fishes = await fetchFishesByBatchIds(batches.map((batch) => batch.id))
  const predictions = await fetchPredictionsByFishIds(fishes.map((fish) => fish.id))

  const images = await fetchImagesByIds([
    ...fishes.map((fish) => fish.eye_image_id),
    ...fishes.map((fish) => fish.gill_image_id)
  ])

  const userMap = new Map(users.map((user) => [user.id, user]))
  const batchMap = new Map(batches.map((batch) => [batch.id, batch]))
  const predictionByFish = new Map(
    predictions.map((prediction) => [prediction.fish_id, prediction])
  )
  const imageMap = new Map(images.map((image) => [image.id, image]))
  const batchesWithStats = buildBatchStats(batches, fishes, predictions)

  if (jenis === 'batch' || jenis === 'batches') {
    const rows = batchesWithStats
      .map((batch) => {
        const user = userMap.get(batch.user_id)

        return {
          'ID Batch': batch.id,
          'Nama User': user?.name || user?.email || 'Tidak diketahui',
          'Email User': user?.email || '-',
          'Jumlah Ikan': batch.fish_count,
          'Kategori Ikan': batch.fish_category || '-',
          Status: batch.status,
          'Status Label': normalizeBatchStatus(batch.status),
          'Confidence Rata-rata': batch.confidenceScoreAvg ?? '-',
          Lokasi: batch.lokasi,
          'Berat Total': batch.berat_total,
          Catatan: batch.catatan || '-',
          'Preprocessed Status': batch.preprocessed_status,
          'Total Inspeksi': batch.totalInspeksi,
          'Total Prediksi': batch.totalPrediksi,
          'Grade A (%)': batch.gradeAPercent,
          'Reject (%)': batch.rejectPercent,
          'Tanggal Dibuat': formatDateTime(batch.created_at),
          'Tanggal Submit': formatDateTime(batch.submitted_at)
        }
      })
      .filter((row) => {
        return (
          matchStatusValue(row.Status, statusFilter) ||
          matchStatusValue(row['Status Label'], statusFilter)
        )
      })
      .map((row, index) => ({ No: index + 1, ...row }))

    return {
      buffer: createWorkbookBuffer([{ name: 'Rekap Batch', rows }]),
      fileName: `rekap-batch-${dari}-sampai-${sampai}.xlsx`
    }
  }

  const rows = fishes
    .map((fish) => {
      const batch = batchMap.get(fish.batch_id)
      const user = batch ? userMap.get(batch.user_id) : null
      const prediction = predictionByFish.get(fish.id)
      const eyeImage = imageMap.get(fish.eye_image_id)
      const gillImage = imageMap.get(fish.gill_image_id)
      const inspectionStatus = getInspectionStatus(fish, prediction)

      return {
        'ID Ikan': fish.id,
        'ID Batch': fish.batch_id,
        'Index Ikan': fish.fish_index,
        'Nama User': user?.name || user?.email || 'Tidak diketahui',
        'Email User': user?.email || '-',
        'Kategori Ikan': batch?.fish_category || '-',
        Lokasi: batch?.lokasi || '-',
        'Status Ikan': fish.status,
        'Status Inspeksi': inspectionStatus,
        Grade: normalizeGrade(prediction?.grade) || '-',
        Label: prediction?.label_text || '-',
        Exportable:
          prediction?.exportable === true
            ? 'Ya'
            : prediction?.exportable === false
              ? 'Tidak'
              : '-',
        'Confidence Score': prediction?.confidence_score ?? '-',
        'Status Mata': prediction?.eyes_status || '-',
        'Confidence Mata': prediction?.eyes_confidence_score ?? '-',
        'Status Insang': prediction?.gill_status || '-',
        'Confidence Insang': prediction?.gill_confidence_score ?? '-',
        'Waktu Proses (ms)': prediction?.waktu_proses_ms ?? '-',
        'File Mata': eyeImage?.file_name || '-',
        'File Insang': gillImage?.file_name || '-',
        'Tanggal Ikan Dibuat': formatDateTime(fish.created_at),
        'Tanggal Prediksi': formatDateTime(prediction?.predicted_at),
        'Tanggal Batch': formatDateTime(batch?.created_at)
      }
    })
    .filter((row) => {
      if (!statusFilter || statusFilter === 'semua') return true

      return (
        matchStatusValue(row['Status Ikan'], statusFilter) ||
        matchStatusValue(row['Status Inspeksi'], statusFilter)
      )
    })
    .map((row, index) => ({ No: index + 1, ...row }))

  return {
    buffer: createWorkbookBuffer([{ name: 'Rekap Inspeksi', rows }]),
    fileName: `rekap-inspeksi-${dari}-sampai-${sampai}.xlsx`
  }
}
