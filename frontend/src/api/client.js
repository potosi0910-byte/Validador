import axios from 'axios'

const api = axios.create({ baseURL: '/api' })

// ── Token en cada petición ────────────────────────────────────────────────────
api.interceptors.request.use(config => {
  const token = localStorage.getItem('drf_token')
  if (token) config.headers.Authorization = `Bearer ${token}`
  return config
})

// ── Redirigir al login si el token expiró ────────────────────────────────────
api.interceptors.response.use(
  res => res,
  err => {
    if (err?.response?.status === 401) {
      localStorage.removeItem('drf_token')
      localStorage.removeItem('drf_user')
      window.location.reload()
    }
    return Promise.reject(err)
  }
)

// ── Reintento automático ante error de red ────────────────────────────────────
// En desarrollo (Vite proxy → backend local en Windows) la primera conexión de
// un socket nuevo a veces se reinicia (ECONNRESET) y el intento siguiente ya
// funciona porque reutiliza una conexión viva. Esto golpea sobre todo al login,
// que suele ser la primerísima petición de la sesión. En vez de obligar al
// usuario a hacer clic varias veces, reintentamos automáticamente (hasta 3
// intentos, con espera creciente) cuando la petición no obtuvo ninguna
// respuesta (fallo de red puro, no un error de negocio 4xx/5xx del servidor).
const MAX_REINTENTOS_RED = 3
api.interceptors.response.use(
  res => res,
  async err => {
    const config = err.config
    const esErrorDeRed = !err.response && err.code !== 'ERR_CANCELED'
    const intento = config?._intentosRed || 0
    if (config && esErrorDeRed && intento < MAX_REINTENTOS_RED) {
      config._intentosRed = intento + 1
      await new Promise(r => setTimeout(r, 300 * config._intentosRed))
      return api(config)
    }
    return Promise.reject(err)
  }
)

// ── Precalentar la conexión con el backend ────────────────────────────────────
// Dispara una petición ligera apenas carga la app, para que cuando el usuario
// llegue a hacer clic en "Ingresar" el socket hacia el backend ya esté
// establecido (evita que el login sea quien sufra el primer fallo de red).
export function precalentarConexion() {
  api.get('/health').catch(() => {})
}

// ── Auth ──────────────────────────────────────────────────────────────────────
export async function login(username, password) {
  const form = new URLSearchParams()
  form.append('username', username)
  form.append('password', password)
  const { data } = await api.post('/auth/login', form, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  })
  localStorage.setItem('drf_token', data.access_token)
  localStorage.setItem('drf_user',  JSON.stringify({
    username: data.username,
    role:     data.role,
    nombre:   data.nombre,
  }))
  return data
}

export function logout() {
  localStorage.removeItem('drf_token')
  localStorage.removeItem('drf_user')
}

export function getUsuarioGuardado() {
  try {
    const raw = localStorage.getItem('drf_user')
    return raw ? JSON.parse(raw) : null
  } catch {
    return null
  }
}

// ── Validación RIPS ───────────────────────────────────────────────────────────
export async function procesar(jsonFiles, excelFiles = [], onProgress) {
  const form = new FormData()
  jsonFiles.forEach(f => form.append('json_files', f))

  // Leer Excel a memoria antes de enviar: evita ERR_UPLOAD_FILE_CHANGED
  // cuando el archivo está abierto en Excel/LibreOffice durante el upload
  for (const f of excelFiles) {
    try {
      const buf = await f.arrayBuffer()
      form.append('excel_files', new Blob([buf], { type: f.type || 'application/octet-stream' }), f.name)
    } catch {
      form.append('excel_files', f)
    }
  }

  const { data } = await api.post('/procesar', form, {
    headers: { 'Content-Type': 'multipart/form-data' },
    timeout: 600000,
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

// ── Administración de usuarios ────────────────────────────────────────────────

export async function listarUsuarios() {
  const { data } = await api.get('/auth/usuarios')
  return data
}

export async function crearUsuario({ username, password, role, nombre }) {
  const { data } = await api.post('/auth/usuarios', { username, password, role, nombre })
  return data
}

export async function eliminarUsuario(username) {
  const { data } = await api.delete(`/auth/usuarios/${username}`)
  return data
}

export async function cambiarPasswordUsuario(username, nueva_password) {
  const { data } = await api.put('/auth/password', { username, nueva_password })
  return data
}

export async function cambiarMiPassword(password_actual, nueva_password) {
  const { data } = await api.put('/auth/me/password', { password_actual, nueva_password })
  return data
}

// ── Machine Learning ──────────────────────────────────────────────────────────
export async function analizarML(excelFile, k = 4) {
  const form = new FormData()
  form.append('excel_file', excelFile)
  form.append('k', String(k))
  const { data } = await api.post('/ml/analizar', form, {
    headers: { 'Content-Type': 'multipart/form-data' },
    timeout: 900000,
  })
  return data
}

export default api
