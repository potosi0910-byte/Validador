import { useEffect, useState } from 'react'

const IcoHSLV = () => (
  <svg viewBox="0 0 24 24" fill="currentColor">
    <path d="M19 3H5C3.9 3 3 3.9 3 5v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10h-4v4h-2v-4H7v-2h4V7h2v4h4v2z"/>
  </svg>
)

const FASES = [
  'Iniciando motor de validación IA...',
  'Leyendo estructura RIPS...',
  'Ejecutando Malla 2275...',
  'Validando autorizaciones EPS...',
  'Ejecutando auditoría clínica...',
  'Generando resultados...',
]

export default function LoadingOverlay({ pct, nFiles }) {
  const [fase, setFase] = useState(0)
  const [elapsed, setElapsed] = useState(0)
  const [mod1, setMod1] = useState(false)
  const [mod2, setMod2] = useState(false)
  const [mod3, setMod3] = useState(false)

  useEffect(() => {
    const t = setInterval(() => setElapsed(s => s + 1), 1000)
    return () => clearInterval(t)
  }, [])

  useEffect(() => {
    const t = setInterval(() => setFase(f => (f + 1) % FASES.length), 1800)
    return () => clearInterval(t)
  }, [])

  useEffect(() => {
    if (pct > 30) setMod1(true)
    if (pct > 60) setMod2(true)
    if (pct > 90) setMod3(true)
  }, [pct])

  const barPct = Math.max(pct, fase * 15)

  return (
    <div className="loading-overlay" style={{ display: 'flex' }}>
      <div className="loading-panel">
        <div className="loading-cross-wrap">
          <div className="lc-ring lc-ring-1"></div>
          <div className="lc-ring lc-ring-2"></div>
          <div className="lc-cross"><IcoHSLV /></div>
        </div>

        <div className="loading-title">Analizando archivos RIPS</div>
        <div className="loading-sub">{FASES[fase]}</div>

        <div className="loading-pct-wrap">
          <div className="loading-pct-track">
            <div className="loading-pct-fill" style={{ width: `${Math.min(barPct, 99)}%` }}></div>
          </div>
          <div className="loading-pct-num">{Math.min(barPct, 99)}%</div>
        </div>

        <div className="loading-counters">
          <div className="lc-stat">
            <span className="lc-num">{nFiles}</span>
            <span className="lc-lbl">archivos JSON</span>
          </div>
          <div className="lc-divider"></div>
          <div className="lc-stat">
            <span className="lc-num">{elapsed}s</span>
            <span className="lc-lbl">transcurrido</span>
          </div>
          <div className="lc-divider"></div>
          <div className="lc-stat">
            <span className="lc-num">110</span>
            <span className="lc-lbl">reglas activas</span>
          </div>
        </div>

        <div className="loading-modules">
          <div className={`lm-item ${mod1 ? 'lm-done' : ''}`}>
            <span className="lm-dot"></span> Malla RIPS 2275
          </div>
          <div className={`lm-item ${mod2 ? 'lm-done' : ''}`}>
            <span className="lm-dot"></span> Validación General
          </div>
          <div className={`lm-item ${mod3 ? 'lm-done' : ''}`}>
            <span className="lm-dot"></span> Auditoría Clínica
          </div>
        </div>
      </div>
    </div>
  )
}
