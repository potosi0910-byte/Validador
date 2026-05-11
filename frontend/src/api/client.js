import axios from 'axios'

const api = axios.create({ baseURL: '/api' })

// ── Token en cada petición ────────────────────────────────────────────────────
api.interceptors.request.use(config => {
  const token = localStorage.getItem('hslv_token')
  if (token) config.headers.Authorization = `Bearer ${token}`
  return config
})

// ── Redirigir al login si el token expiró ────────────────────────────────────
api.interceptors.response.use(
  res => res,
  err => {
    if (err?.response?.status === 401) {
      localStorage.removeItem('hslv_token')
      localStorage.removeItem('hslv_user')
      window.location.reload()
    }
    return Promise.reject(err)
  }
)

// ── Auth ──────────────────────────────────────────────────────────────────────
export async function login(username, password) {
  const form = new URLSearchParams()
  form.append('username', username)
  form.append('password', password)
  const { data } = await api.post('/auth/login', form, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  })
  localStorage.setItem('hslv_token', data.access_token)
  localStorage.setItem('hslv_user',  JSON.stringify({
    username: data.username,
    role:     data.role,
    nombre:   data.nombre,
  }))
  return data
}

export function logout() {
  localStorage.removeItem('hslv_token')
  localStorage.removeItem('hslv_user')
}

export function getUsuarioGuardado() {
  try {
    const raw = localStorage.getItem('hslv_user')
    return raw ? JSON.parse(raw) : null
  } catch {
    return null
  }
}

// ── Validación RIPS ───────────────────────────────────────────────────────────
export async function procesar(jsonFiles, excelFiles = [], onProgress) {
  const form = new FormData()
  jsonFiles.forEach(f  => form.append('json_files',   f))
  excelFiles.forEach(f => form.append('excel_files',  f))

  const { data } = await api.post('/procesar', form, {
    headers: { 'Content-Type': 'multipart/form-data' },
    onUploadProgress: e => {
      if (onProgress && e.total) onProgress(Math.round((e.loaded * 100) / e.total))
    },
  })
  return data
}

export async function exportarExcel() {
  const resp = await api.get('/exportar', { responseType: 'blob' })
  const url  = URL.createObjectURL(resp.data)
  const a    = document.createElement('a')
  const cd   = resp.headers['content-disposition'] || ''
  const match = cd.match(/filename=([^;]+)/)
  a.href     = url
  a.download = match ? match[1].replace(/"/g, '') : 'Alertas_Malla_Validadora.xlsx'
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

export default api
