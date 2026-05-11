import { useState } from 'react'
import { login } from '../api/client'

const IcoHSLV = () => (
  <svg viewBox="0 0 24 24" fill="currentColor">
    <path d="M19 3H5C3.9 3 3 3.9 3 5v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10h-4v4h-2v-4H7v-2h4V7h2v4h4v2z"/>
  </svg>
)

const IcoUser = () => (
  <svg viewBox="0 0 20 20" fill="currentColor">
    <path fillRule="evenodd" d="M10 9a3 3 0 100-6 3 3 0 000 6zm-7 9a7 7 0 1114 0H3z" clipRule="evenodd"/>
  </svg>
)

const IcoLock = () => (
  <svg viewBox="0 0 20 20" fill="currentColor">
    <path fillRule="evenodd" d="M5 9V7a5 5 0 0110 0v2a2 2 0 012 2v5a2 2 0 01-2 2H5a2 2 0 01-2-2v-5a2 2 0 012-2zm8-2v2H7V7a3 3 0 016 0z" clipRule="evenodd"/>
  </svg>
)

export default function LoginPage({ onLogin }) {
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError]       = useState('')
  const [loading, setLoading]   = useState(false)

  const handleSubmit = async (e) => {
    e.preventDefault()
    if (!username || !password) {
      setError('Ingrese usuario y contraseña.')
      return
    }
    setLoading(true)
    setError('')
    try {
      const data = await login(username, password)
      onLogin(data)
    } catch (err) {
      setError(err?.response?.data?.detail ?? 'Usuario o contraseña incorrectos.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <>
      {/* Fondo */}
      <div className="bg-glow-top"></div>
      <div className="bg-grid"></div>
      <div className="bg-grid-floor"></div>
      <div className="bg-cross bg-cross-1"><IcoHSLV /></div>
      <div className="bg-cross bg-cross-2"><IcoHSLV /></div>
      <div className="bg-cross bg-cross-3"><IcoHSLV /></div>
      <div className="orb orb-1"></div>
      <div className="orb orb-2"></div>
      <div className="orb orb-3"></div>

      <div className="shell" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: '100vh' }}>

        {/* Logo y título */}
        <div style={{ textAlign: 'center', marginBottom: '2rem' }}>
          <div className="brand-icon" style={{ width: '4rem', height: '4rem', fontSize: '2rem', margin: '0 auto 1rem' }}>
            <IcoHSLV />
          </div>
          <div className="brand-name" style={{ fontSize: '1.25rem', fontWeight: 700 }}>
            Hospital Susana López de Valencia
          </div>
          <div className="brand-sub">Plataforma de Validación RIPS &amp; Auditoría Clínica</div>
        </div>

        {/* Card de login */}
        <div className="card upload-card" style={{ maxWidth: '400px', width: '100%' }}>
          <div className="card-label">
            <IcoLock />
            Acceso al sistema
          </div>

          <form onSubmit={handleSubmit} style={{ padding: '1.5rem', display: 'flex', flexDirection: 'column', gap: '1rem' }}>

            {/* Usuario */}
            <div className="login-field">
              <label className="login-label">Usuario</label>
              <div className="login-input-wrap">
                <span className="login-ico"><IcoUser /></span>
                <input
                  className="login-input"
                  type="text"
                  placeholder="Nombre de usuario"
                  value={username}
                  onChange={e => setUsername(e.target.value)}
                  autoComplete="username"
                  autoFocus
                  disabled={loading}
                />
              </div>
            </div>

            {/* Contraseña */}
            <div className="login-field">
              <label className="login-label">Contraseña</label>
              <div className="login-input-wrap">
                <span className="login-ico"><IcoLock /></span>
                <input
                  className="login-input"
                  type="password"
                  placeholder="Contraseña"
                  value={password}
                  onChange={e => setPassword(e.target.value)}
                  autoComplete="current-password"
                  disabled={loading}
                />
              </div>
            </div>

            {/* Error */}
            {error && (
              <div className="notice-jorge" style={{ margin: 0 }}>
                <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '1rem', height: '1rem', flexShrink: 0 }}>
                  <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7 4a1 1 0 11-2 0 1 1 0 012 0zm-1-9a1 1 0 00-1 1v4a1 1 0 102 0V6a1 1 0 00-1-1z" clipRule="evenodd"/>
                </svg>
                {error}
              </div>
            )}

            {/* Botón */}
            <button
              type="submit"
              className="btn-primary"
              disabled={loading}
              style={{ marginTop: '0.5rem', justifyContent: 'center' }}
            >
              {loading ? (
                <>
                  <span className="btn-cargando-spin" style={{ display: 'inline-block' }}></span>
                  Verificando...
                </>
              ) : (
                <>
                  <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '1rem', height: '1rem' }}>
                    <path fillRule="evenodd" d="M3 3a1 1 0 011 1v12a1 1 0 11-2 0V4a1 1 0 011-1zm7.707 3.293a1 1 0 010 1.414L9.414 9H17a1 1 0 110 2H9.414l1.293 1.293a1 1 0 01-1.414 1.414l-3-3a1 1 0 010-1.414l3-3a1 1 0 011.414 0z" clipRule="evenodd"/>
                  </svg>
                  Ingresar
                </>
              )}
            </button>

          </form>
        </div>

        <div style={{ marginTop: '1.5rem', opacity: 0.5, fontSize: '0.75rem', textAlign: 'center', color: 'var(--text-muted, #94a3b8)' }}>
          Resoluciones 2275 y 2284 de 2023 &nbsp;·&nbsp; Desarrollado por Didier Potosi
        </div>

      </div>
    </>
  )
}
