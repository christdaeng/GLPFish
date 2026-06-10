import api from '../lib/api'

export const AdminService = {
  async getSummary(params = {}) {
    const { data } = await api.get('/admin/summary', { params })
    return data.summary
  },

  async getChart(params = {}) {
    const { data } = await api.get('/admin/chart', { params })
    return data.chart
  },

  async getTopUsers(params = {}) {
    const { data } = await api.get('/admin/top-users', { params })
    return data.users
  },

  async getRecentBatches(params = {}) {
    const { data } = await api.get('/admin/recent-batches', { params })
    return data.batches
  },

  async getActivity(params = {}) {
    const { data } = await api.get('/admin/activity', { params })
    return data.activities
  },

  async getUsers(params = {}) {
    const { data } = await api.get('/admin/users', { params })
    return data
  },

  async getUserDetail(id) {
    const { data } = await api.get(`/admin/users/${id}`)
    return data.user
  },

  async downloadRecap(params = {}) {
    const response = await api.get('/admin/recap', {
      params,
      responseType: 'blob'
    })

    return response.data
  },

  async saveRecap(params = {}, filename = 'rekap-admin.xlsx') {
    const blob = await this.downloadRecap(params)
    const url = window.URL.createObjectURL(blob)
    const link = document.createElement('a')

    link.href = url
    link.download = filename
    document.body.appendChild(link)
    link.click()
    link.remove()
    window.URL.revokeObjectURL(url)
  }
}

export default AdminService
