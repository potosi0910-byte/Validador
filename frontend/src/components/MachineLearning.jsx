// MachineLearning.jsx — Dashboard Analítico · Estancia Hospitalaria
// Renderiza dashboard_ml.html (Chart.js + xlsx.js) en iframe full-height.
// Todo el procesamiento ocurre en el browser: carga inmediata, sin servidor.

import { useState, useRef } from 'react'

const IcoBrain = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" style={{width:26,height:26}}>
    <path strokeLinecap="round" d="M9.5 2A2.5 2.5 0 007 4.5v.5a2.5 2.5 0 00-2.5 2.5 2.5 2.5 0 00-2.5 2.5 2.5 2.5 0 002.5 2.5 2.5 2.5 0 002.5 2.5v.5a2.5 2.5 0 002.5 2.5"/>
    <path strokeLinecap="round" d="M14.5 2A2.5 2.5 0 0117 4.5v.5a2.5 2.5 0 012.5 2.5 2.5 2.5 0 012.5 2.5 2.5 2.5 0 01-2.5 2.5 2.5 2.5 0 01-2.5 2.5v.5a2.5 2.5 0 01-2.5 2.5"/>
    <line x1="12" y1="6" x2="12" y2="18" strokeLinecap="round"/>
    <line x1="9" y1="10" x2="15" y2="10" strokeLinecap="round"/>
    <line x1="9" y1="14" x2="15" y2="14" strokeLinecap="round"/>
  </svg>
)

const IcoExpand = () => (
  <svg viewBox="0 0 20 20" fill="currentColor" style={{width:15,height:15}}>
    <path fillRule="evenodd" d="M3 4a1 1 0 011-1h4a1 1 0 010 2H6.414l2.293 2.293a1 1 0 01-1.414 1.414L5 6.414V8a1 1 0 01-2 0V4zm9 1a1 1 0 010-2h4a1 1 0 011 1v4a1 1 0 01-2 0V6.414l-2.293 2.293a1 1 0 11-1.414-1.414L13.586 5H12zm-9 7a1 1 0 012 0v1.586l2.293-2.293a1 1 0 011.414 1.414L6.414 15H8a1 1 0 010 2H4a1 1 0 01-1-1v-4zm13-1a1 1 0 011 1v4a1 1 0 01-1 1h-4a1 1 0 010-2h1.586l-2.293-2.293a1 1 0 011.414-1.414L15 13.586V12a1 1 0 011-1z" clipRule="evenodd"/>
  </svg>
)

const IcoCollapse = () => (
  <svg viewBox="0 0 20 20" fill="currentColor" style={{width:15,height:15}}>
    <path fillRule="evenodd" d="M5 10a1 1 0 011-1h8a1 1 0 110 2H6a1 1 0 01-1-1z" clipRule="evenodd"/>
  </svg>
)

export default function MachineLearning() {
  const [loaded, setLoaded]       = useState(false)
  const [fullscreen, setFullscreen] = useState(false)
  const iframeRef = useRef(null)

  return (
    <div className={`ml-wrap${fullscreen ? ' ml-wrap-full' : ''}`}>

      {/* ── Barra de título ─────────────────────────────────────────── */}
      <div className="ml-topbar">
        <div className="ml-topbar-left">
          <div className="ml-topbar-icon"><IcoBrain /></div>
          <div>
            <div className="ml-topbar-title">Dashboard Hospitalario · Análisis Integral</div>
            <div className="ml-topbar-sub">
              Hospitalización · UCI · Cirugía · Ambulatorio · Financiero · Predictor Clínico
            </div>
          </div>
        </div>
        <div className="ml-topbar-badges">
          <span className="ml-badge ml-badge-cyan">Chart.js</span>
          <span className="ml-badge ml-badge-blue">SheetJS</span>
          <span className="ml-badge ml-badge-purple">K-Means · Predictor CIE-10</span>
          <button
            className="ml-expand-btn"
            onClick={() => setFullscreen(f => !f)}
            title={fullscreen ? 'Contraer' : 'Pantalla completa'}
          >
            {fullscreen ? <IcoCollapse /> : <IcoExpand />}
            {fullscreen ? 'Contraer' : 'Expandir'}
          </button>
        </div>
      </div>

      {/* ── iframe del dashboard ─────────────────────────────────────── */}
      <div className="ml-iframe-shell">
        {!loaded && (
          <div className="ml-iframe-loading">
            <div className="ml-loading-anim">
              <div className="ml-loading-ring ml-loading-ring-1"></div>
              <div className="ml-loading-ring ml-loading-ring-2"></div>
              <div className="ml-loading-ring ml-loading-ring-3"></div>
              <div className="ml-loading-core"><IcoBrain /></div>
            </div>
            <div className="ml-loading-text">Cargando módulo de análisis...</div>
          </div>
        )}
        <iframe
          ref={iframeRef}
          src="/dashboard_ml.html"
          title="Dashboard ML · Análisis Estancia Hospitalaria"
          className={`ml-iframe${loaded ? ' ml-iframe-visible' : ''}`}
          onLoad={() => setLoaded(true)}
          allow="clipboard-write"
        />
      </div>

    </div>
  )
}
