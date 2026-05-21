// Sidebar.jsx — Navegación lateral izquierda

const IcoShield = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6">
    <path strokeLinecap="round" strokeLinejoin="round"
      d="M12 2L3 6v6c0 5.25 3.75 10.15 9 11.25C17.25 22.15 21 17.25 21 12V6L12 2z"/>
    <path strokeLinecap="round" d="M8.5 12l2.5 2.5 4.5-4.5"/>
  </svg>
)

const IcoBrain = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6">
    <path strokeLinecap="round" d="M9.5 2A2.5 2.5 0 007 4.5v.5a2.5 2.5 0 00-2.5 2.5 2.5 2.5 0 00-2.5 2.5 2.5 2.5 0 002.5 2.5 2.5 2.5 0 002.5 2.5v.5a2.5 2.5 0 002.5 2.5"/>
    <path strokeLinecap="round" d="M14.5 2A2.5 2.5 0 0117 4.5v.5a2.5 2.5 0 012.5 2.5 2.5 2.5 0 012.5 2.5 2.5 2.5 0 01-2.5 2.5 2.5 2.5 0 01-2.5 2.5v.5a2.5 2.5 0 01-2.5 2.5"/>
    <line x1="12" y1="6" x2="12" y2="18" strokeLinecap="round"/>
    <line x1="9" y1="10" x2="15" y2="10" strokeLinecap="round"/>
    <line x1="9" y1="14" x2="15" y2="14" strokeLinecap="round"/>
  </svg>
)

const IcoRips = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6">
    <path strokeLinecap="round" d="M9 12h6M9 16h4M13 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V9z"/>
    <path strokeLinecap="round" d="M13 2v7h7"/>
  </svg>
)

const IcoCheck = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6">
    <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/>
  </svg>
)

const IcoGear = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6">
    <path strokeLinecap="round" strokeLinejoin="round" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z"/>
    <path strokeLinecap="round" strokeLinejoin="round" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/>
  </svg>
)

const ALL_SECTIONS = [
  {
    id: 'malla',
    label: 'Malla Validadora',
    desc: 'RIPS · Res. 2275/2023',
    icon: <IcoShield />,
    badgeColor: 'brand',
    roles: null,
  },
  {
    id: 'ml',
    label: 'Machine Learning',
    desc: 'Análisis Predictivo · IA',
    icon: <IcoBrain />,
    badgeColor: 'accent',
    roles: null,
  },
  {
    id: 'admin',
    label: 'Administración',
    desc: 'Usuarios · Seguridad',
    icon: <IcoGear />,
    badgeColor: 'warn',
    roles: ['admin'],
  },
]

export default function Sidebar({ active, onSelect, usuario }) {
  const SECTIONS = ALL_SECTIONS.filter(s => !s.roles || s.roles.includes(usuario?.role))
  return (
    <aside className="sidebar">
      <div className="sidebar-head">
        <span className="sidebar-head-label">Módulos</span>
      </div>
      <nav className="sidebar-nav">
        {SECTIONS.map(s => (
          <button
            key={s.id}
            className={`sidebar-item${active === s.id ? ' sidebar-item-active' : ''}`}
            onClick={() => onSelect(s.id)}
            title={s.label}
          >
            <div className={`sidebar-ico sidebar-ico-${s.badgeColor}`}>
              {s.icon}
            </div>
            <div className="sidebar-text">
              <span className="sidebar-label">{s.label}</span>
              <span className="sidebar-desc">{s.desc}</span>
            </div>
            {active === s.id && <div className="sidebar-active-bar" />}
          </button>
        ))}
      </nav>

      <div className="sidebar-footer">
        <div className="sidebar-version">
          <IcoCheck />
          v2.0 · Dr. Factory
        </div>
      </div>
    </aside>
  )
}
