import { useState, useEffect } from 'react'
import LoadingOverlay   from './components/LoadingOverlay'
import LoginPage        from './components/LoginPage'
import UploadCard       from './components/UploadCard'
import ResultsSection   from './components/ResultsSection'
import Sidebar          from './components/Sidebar'
import MachineLearning  from './components/MachineLearning'
import Admin            from './components/Admin'
import { procesar, exportarExcel, logout, getUsuarioGuardado } from './api/client'

const IcoFactory = () => (
  <svg viewBox="0 0 24 24" fill="currentColor">
    <path d="M19 3H5C3.9 3 3 3.9 3 5v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10h-4v4h-2v-4H7v-2h4V7h2v4h4v2z"/>
  </svg>
)

const ROLE_LABEL = { admin: 'Administrador', auditor: 'Auditor Clínico' }

export default function App() {
  const [usuario, setUsuario]       = useState(null)
  const [section, setSection]       = useState('malla')
  const [loading, setLoading]       = useState(false)
  const [loadingPct, setLoadingPct] = useState(0)
  const [resultado, setResultado]   = useState(null)
  const [error, setError]           = useState(null)
  const [jsonFiles, setJsonFiles]   = useState([])
  const [excelFiles, setExcelFiles] = useState([])

  useEffect(() => {
    const u = getUsuarioGuardado()
    if (u) setUsuario(u)
  }, [])

  const handleLogin = (data) => {
    setUsuario({ username: data.username, role: data.role, nombre: data.nombre })
  }

  const handleLogout = () => {
    logout()
    setUsuario(null)
    setResultado(null)
    setError(null)
    setJsonFiles([])
    setExcelFiles([])
    setSection('malla')
  }

  const handleProcesar = async () => {
    if (!jsonFiles.length) return
    setLoading(true)
    setLoadingPct(0)
    setError(null)
    try {
      const data = await procesar(jsonFiles, excelFiles, pct => setLoadingPct(pct))
      setResultado(data)
    } catch (err) {
      setError(err?.response?.data?.detail ?? err?.message ?? 'Error al procesar los archivos.')
    } finally {
      setLoading(false)
    }
  }

  const handleExportar = async (e) => {
    e.preventDefault()
    try {
      await exportarExcel()
    } catch {
      setError('No se pudo descargar el Excel.')
    }
  }

  if (!usuario) {
    return <LoginPage onLogin={handleLogin} />
  }

  const hasResults = !!resultado

  return (
    <>
      {/* Fondo 3D — compartido por todas las secciones */}
      <div className="bg-glow-top"></div>
      <div className="bg-grid"></div>
      <div className="bg-grid-floor"></div>
      <div className="bg-cross bg-cross-1"><IcoFactory /></div>
      <div className="bg-cross bg-cross-2"><IcoFactory /></div>
      <div className="bg-cross bg-cross-3"><IcoFactory /></div>
      <div className="orb orb-1"></div>
      <div className="orb orb-2"></div>
      <div className="orb orb-3"></div>

      {loading && <LoadingOverlay pct={loadingPct} nFiles={jsonFiles.length} />}

      <div className="shell">

        {/* Header */}
        <header className="topbar">
          <div className="topbar-brand">
            <div className="brand-icon"><IcoFactory /></div>
            <div>
              <span className="brand-name">Dr. Factory</span>
              <span className="brand-sub">Plataforma de Validación RIPS &amp; Auditoría Clínica</span>
            </div>
          </div>
          <div className="topbar-right">
            <div className="topbar-badge">
              <span className="pulse-dot"></span>
              Res. 2275 / 2023
            </div>
            <div className="topbar-badge topbar-badge-ai">
              <svg viewBox="0 0 20 20" fill="currentColor">
                <path d="M11 3a1 1 0 10-2 0v1a1 1 0 102 0V3zM15.657 5.757a1 1 0 00-1.414-1.414l-.707.707a1 1 0 001.414 1.414l.707-.707zM18 10a1 1 0 01-1 1h-1a1 1 0 110-2h1a1 1 0 011 1zM5.05 6.464A1 1 0 106.464 5.05l-.707-.707a1 1 0 00-1.414 1.414l.707.707zM5 10a1 1 0 01-1 1H3a1 1 0 110-2h1a1 1 0 011 1zm3 6v-1h4v1a2 2 0 11-4 0zm4-2a4 4 0 10-4 0h4z"/>
              </svg>
              IA Activa
            </div>
            <div className="topbar-badge topbar-user">
              <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '0.9rem', height: '0.9rem' }}>
                <path fillRule="evenodd" d="M10 9a3 3 0 100-6 3 3 0 000 6zm-7 9a7 7 0 1114 0H3z" clipRule="evenodd"/>
              </svg>
              <span>{usuario.nombre || usuario.username}</span>
              <span className="user-role-tag">{ROLE_LABEL[usuario.role] ?? usuario.role}</span>
            </div>
            <button className="btn-logout" onClick={handleLogout} title="Cerrar sesión">
              <svg viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M3 3a1 1 0 00-1 1v12a1 1 0 102 0V4a1 1 0 00-1-1zm10.293 9.293a1 1 0 001.414 1.414l3-3a1 1 0 000-1.414l-3-3a1 1 0 10-1.414 1.414L14.586 9H7a1 1 0 100 2h7.586l-1.293 1.293z" clipRule="evenodd"/>
              </svg>
              Salir
            </button>
          </div>
        </header>

        {/* Cuerpo: sidebar izquierdo + contenido principal */}
        <div className="app-body">

          <Sidebar active={section} onSelect={setSection} usuario={usuario} />

          <main className="main-content">

            {/* ── Sección: Malla Validadora RIPS ── */}
            {section === 'malla' && (
              <>
                <UploadCard
                  jsonFiles={jsonFiles}
                  excelFiles={excelFiles}
                  onJsonChange={setJsonFiles}
                  onExcelChange={setExcelFiles}
                  onProcesar={handleProcesar}
                  onExportar={handleExportar}
                  hasResults={hasResults}
                  error={error}
                  loading={loading}
                />
                {hasResults && <ResultsSection resultado={resultado} />}
              </>
            )}

            {/* ── Sección: Machine Learning ── */}
            {section === 'ml' && <MachineLearning />}

            {/* ── Sección: Administración ── */}
            {section === 'admin' && <Admin usuario={usuario} />}

          </main>
        </div>

        <footer className="footer">
          Dr Factory — Sistema de Validación RIPS &amp; Auditoría Clínica &nbsp;·&nbsp; Resoluciones 2275 y 2284 de 2023
          <br />
          <small>Desarrollado por Didier Potosi derechos reservados 2026</small>
        </footer>

      </div>
    </>
  )
}
