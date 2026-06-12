import { useState } from 'react'
import { login } from '../api/client'

// ── Íconos ────────────────────────────────────────────────────────────────────
const IcoUser = () => (
  <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.6">
    <path strokeLinecap="round" d="M10 10a3.5 3.5 0 100-7 3.5 3.5 0 000 7zm-6.5 7.5a6.5 6.5 0 0113 0"/>
  </svg>
)
const IcoLock = () => (
  <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.6">
    <rect x="4" y="9" width="12" height="9" rx="2"/>
    <path strokeLinecap="round" d="M7 9V6a3 3 0 016 0v3"/>
  </svg>
)
const IcoEye = ({ open }) => open
  ? <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.6"><path strokeLinecap="round" d="M1 10s3.6-7 9-7 9 7 9 7-3.6 7-9 7-9-7-9-7z"/><circle cx="10" cy="10" r="2.5"/></svg>
  : <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.6"><path strokeLinecap="round" d="M2 2l16 16M7.5 7.6A3 3 0 0013 13M5.4 5.5C3.3 6.9 1.8 8.8 1 10c1.7 2.8 4.9 7 9 7 1.7 0 3.3-.6 4.6-1.5M8 3.2A8.5 8.5 0 0110 3c4.1 0 7.3 4.2 9 7a16 16 0 01-1.9 2.7"/></svg>
const IcoShield = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
    <path strokeLinecap="round" strokeLinejoin="round" d="M12 2L3 6v6c0 5.25 3.75 10.15 9 11.25C17.25 22.15 21 17.25 21 12V6L12 2z"/>
    <path strokeLinecap="round" d="M8.5 12l2.5 2.5 4.5-4.5"/>
  </svg>
)
const IcoBrain = () => (
  <svg viewBox="0 0 64 64" fill="none" stroke="currentColor" strokeWidth="1.2">
    <ellipse cx="32" cy="32" rx="20" ry="18" strokeOpacity=".6"/>
    <path strokeLinecap="round" d="M22 24c0-5.5 4-9 10-9s10 3.5 10 9"/>
    <path strokeLinecap="round" d="M18 32c-3 0-6 2-6 5s3 5 6 5M46 32c3 0 6 2 6 5s-3 5-6 5"/>
    <path strokeLinecap="round" d="M24 48c1.5 2.5 4 4 8 4s6.5-1.5 8-4"/>
    <circle cx="26" cy="30" r="1.5" fill="currentColor" strokeWidth="0"/>
    <circle cx="38" cy="30" r="1.5" fill="currentColor" strokeWidth="0"/>
    <circle cx="32" cy="36" r="1.5" fill="currentColor" strokeWidth="0"/>
    <line x1="26" y1="30" x2="32" y2="36"/><line x1="38" y1="30" x2="32" y2="36"/>
    <line x1="26" y1="30" x2="38" y2="30"/>
  </svg>
)
const IcoChart = () => <svg viewBox="0 0 20 20" fill="currentColor"><path d="M2 11a1 1 0 011-1h2a1 1 0 011 1v5a1 1 0 01-1 1H3a1 1 0 01-1-1v-5zm6-4a1 1 0 011-1h2a1 1 0 011 1v9a1 1 0 01-1 1H9a1 1 0 01-1-1V7zm6-3a1 1 0 011-1h2a1 1 0 011 1v12a1 1 0 01-1 1h-2a1 1 0 01-1-1V4z"/></svg>
const IcoZap   = () => <svg viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M11.3 1.046A1 1 0 0112 2v5h4a1 1 0 01.82 1.573l-7 10A1 1 0 018 18v-5H4a1 1 0 01-.82-1.573l7-10a1 1 0 011.12-.38z" clipRule="evenodd"/></svg>
const IcoDoc   = () => <svg viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M4 4a2 2 0 012-2h4.586A2 2 0 0112 2.586L15.414 6A2 2 0 0116 7.414V16a2 2 0 01-2 2H6a2 2 0 01-2-2V4z" clipRule="evenodd"/></svg>

// ── Componente principal ──────────────────────────────────────────────────────
export default function LoginPage({ onLogin }) {
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [showPwd, setShowPwd]   = useState(false)
  const [error, setError]       = useState('')
  const [loading, setLoading]   = useState(false)

  const handleSubmit = async (e) => {
    e.preventDefault()
    setLoading(true); setError('')
    try {
      const data = await login(username || 'admin', password || 'x')
      onLogin(data)
    } catch (err) {
      setError(err?.response?.data?.detail ?? 'Error de conexión con el servidor.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="lx-root">

      {/* ══ PANEL IZQUIERDO ═══════════════════════════════════════════════════ */}
      <div className="lx-left">

        {/* Logo */}
        <div className="lx-logo-wrap">
          <div className="lx-logo-ico">
            <svg viewBox="0 0 36 36" fill="none">
              <defs>
                <linearGradient id="lgLogo" x1="0" y1="0" x2="1" y2="1">
                  <stop offset="0%" stopColor="#60a5fa"/>
                  <stop offset="100%" stopColor="#a78bfa"/>
                </linearGradient>
              </defs>
              <polygon points="18,3 33,30 3,30" stroke="url(#lgLogo)" strokeWidth="2.5" fill="none"/>
              <line x1="18" y1="10" x2="18" y2="22" stroke="url(#lgLogo)" strokeWidth="2" strokeLinecap="round"/>
              <circle cx="18" cy="26" r="1.8" fill="url(#lgLogo)"/>
            </svg>
          </div>
          <div>
            <div className="lx-brand">DR. FACTORY</div>
            <div className="lx-brand-sub">INTELIGENCIA QUE VALIDA, CONFIANZA QUE TRANSFORMA</div>
          </div>
        </div>

        {/* Badge */}
        <div className="lx-badge">
          <IcoShield />
          Plataforma de Auditoría Inteligente con IA
        </div>

        {/* Título */}
        <div className="lx-title">
          Bienvenido de <span className="lx-accent">vuelta</span>
        </div>
        <p className="lx-desc">
          Ingresa a tu plataforma y continúa optimizando tus procesos de auditoría con{' '}
          <span className="lx-accent">Inteligencia Artificial</span>.
        </p>

        {/* Formulario */}
        <form onSubmit={handleSubmit} className="lx-form">

          <div className="lx-field">
            <label className="lx-label">Usuario</label>
            <div className="lx-input-wrap">
              <input
                className="lx-input"
                style={{ paddingLeft: '.9rem' }}
                type="text"
                placeholder="Ingresa tu usuario"
                value={username}
                onChange={e => setUsername(e.target.value)}
                autoComplete="username"
                autoFocus
                disabled={loading}
              />
            </div>
          </div>

          <div className="lx-field">
            <label className="lx-label">Contraseña</label>
            <div className="lx-input-wrap">
              <input
                className="lx-input"
                style={{ paddingLeft: '.9rem' }}
                type={showPwd ? 'text' : 'password'}
                placeholder="Ingresa tu contraseña"
                value={password}
                onChange={e => setPassword(e.target.value)}
                autoComplete="current-password"
                disabled={loading}
              />
              <button type="button" className="lx-eye" onClick={() => setShowPwd(v => !v)} tabIndex={-1}>
                <IcoEye open={showPwd} />
              </button>
            </div>
          </div>

          {error && (
            <div className="lx-error">
              <svg viewBox="0 0 20 20" fill="currentColor" style={{width:'1rem',height:'1rem',flexShrink:0}}>
                <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7 4a1 1 0 11-2 0 1 1 0 012 0zm-1-9a1 1 0 00-1 1v4a1 1 0 102 0V6a1 1 0 00-1-1z" clipRule="evenodd"/>
              </svg>
              {error}
            </div>
          )}

          <button type="submit" className="lx-btn" disabled={loading}>
            {loading
              ? <><span className="lx-spin"></span> Verificando...</>
              : <>Iniciar sesión <span className="lx-arrow">→</span></>
            }
          </button>

        </form>

        {/* Seguridad */}
        <div className="lx-security">
          <div className="lx-sec-ico"><IcoShield /></div>
          <div>
            <div className="lx-sec-title">Seguridad &amp; Confianza</div>
            <div className="lx-sec-desc">Encriptación JWT, cumplimiento normativo y protección de datos garantizados.</div>
          </div>
        </div>
        <div className="lx-sec-online">
          <span className="lx-dot-green"></span> Sistema 100% Seguro
        </div>

        <div className="lx-footer-copy">© 2026 Dr. Factory · Desarrollado por Didier Potosi</div>
      </div>

      {/* ══ PANEL DERECHO ═════════════════════════════════════════════════════ */}
      <div className="lx-right">
        <div className="lx-right-inner">

          {/* Cita principal */}
          <div className="lx-quote-top">
            "La auditoría del futuro,<br/>
            impulsada por <span className="lx-accent-r">Inteligencia Artificial</span>"
          </div>

          {/* Animación central IA */}
          <div className="lx-brain-scene">
            <div className="lx-ring lx-ring-3"></div>
            <div className="lx-ring lx-ring-2"></div>
            <div className="lx-ring lx-ring-1"></div>
            <div className="lx-brain-core">
              <div className="lx-brain-glow"></div>
              <div className="lx-brain-ico"><IcoBrain /></div>
              <div className="lx-ai-label">Auditoria</div>
            </div>
            <div className="lx-orbit lx-orbit-1"><div className="lx-orb-dot"></div></div>
            <div className="lx-orbit lx-orbit-2"><div className="lx-orb-dot lx-orb-dot-2"></div></div>
            <div className="lx-orbit lx-orbit-3"><div className="lx-orb-dot lx-orb-dot-3"></div></div>
          </div>

          {/* Stats */}
          <div className="lx-stats-row">
            <div className="lx-stat-card">
              <div className="lx-stat-label">Análisis Inteligente</div>
              <div className="lx-stat-val lx-stat-green">↑ 37% <span>Eficiencia</span></div>
            </div>
            <div className="lx-stat-card">
              <div className="lx-stat-label">Riesgos Detectados</div>
              <div className="lx-stat-val"> Clasificación<span></span></div>
              <div className="lx-stat-dots">
                <span className="sd-red"></span>Críticos
                <span className="sd-yellow"></span>Medios
                <span className="sd-green"></span>Bajos
              </div>
            </div>
            <div className="lx-stat-card">
              <div className="lx-stat-label">Validaciones por cargue</div>
              <div className="lx-stat-val lx-stat-big">900 <span>Rips</span></div>
              <div className="lx-stat-sub">Documentos procesados</div>
            </div>
          </div>

          {/* Features */}
          <div className="lx-features">
            {[
              { ico: <IcoBrain />, label: 'IA Avanzada',          desc: 'Algoritmos de última generación' },
              { ico: <IcoChart />, label: 'Análisis Predictivo',  desc: 'Anticipa riesgos y oportunidades' },
              { ico: <IcoZap />,   label: 'Automatización',        desc: 'Procesos inteligentes 24/7' },
              { ico: <IcoDoc />,   label: 'Reportes Inteligentes', desc: 'Información clara, decisiones precisas' },
            ].map(f => (
              <div className="lx-feat" key={f.label}>
                <div className="lx-feat-ico">{f.ico}</div>
                <div className="lx-feat-label">{f.label}</div>
                <div className="lx-feat-desc">{f.desc}</div>
              </div>
            ))}
          </div>

          {/* Cita inferior */}
          <div className="lx-quote-bottom">
            <span className="lx-quote-mark">"</span>
            No solo detectamos lo que pasó, entendemos por <em>qué pasó</em>{' '}
            y predecimos lo que <em>puede pasar</em>.
            <div className="lx-quote-author">— Auditoría inteligente, decisiones inteligentes.</div>
          </div>

        </div>
      </div>

    </div>
  )
}
