import { useRef, useState, useEffect } from 'react'

const RIPS_PREFIJO = 'Rips_SL'
const RIPS_MAX     = 5000

function fmtFecha(ts) {
  return new Date(ts).toISOString().slice(0, 10)
}

export default function UploadCard({
  jsonFiles, excelFiles,
  onJsonChange, onExcelChange,
  onProcesar, onExportar,
  hasResults, error, loading, loadingPct = 0,
}) {
  const [modo, setModo]               = useState('auto')
  const [filtro, setFiltro]           = useState('fecha')
  const [autoTodos, setAutoTodos]     = useState([])
  const [diasPorFecha, setDiasPorFecha] = useState([])
  const [fechaDesde, setFechaDesde]   = useState('')
  const [fechaHasta, setFechaHasta]   = useState('')
  const [numDesde, setNumDesde]       = useState(1)
  const [mostrarResumen, setMostrarResumen] = useState(false)
  const [clasActivo, setClasActivo]       = useState(false)
  const [btnCargando, setBtnCargando]     = useState(false)
  const [esperandoCarpeta, setEsperandoCarpeta] = useState(false)
  const [escaneando, setEscaneando]       = useState(false)
  const [procesando, setProcesando]       = useState(false)

  useEffect(() => {
    if (!loading) setProcesando(false)
  }, [loading])

  const inputAutoRef      = useRef()
  const inputManualRef    = useRef()
  const inputExcelRef     = useRef()
  const carpetaOkRef      = useRef(false)  // true cuando onChange disparó (carpeta confirmada)

  // ── Aplicar filtro ────────────────────────────────────────────────────────
  function aplicarFiltro(todos, fd, fh, nd, flt) {
    let filtrados = todos
    if (flt === 'fecha') {
      filtrados = todos.filter(f => {
        const dia = fmtFecha(f.lastModified)
        return (!fd || dia >= fd) && (!fh || dia <= fh)
      })
    } else {
      filtrados = todos.slice(0, nd ? (todos.length - nd + 1) : todos.length)
    }
    filtrados = filtrados.slice(0, RIPS_MAX)
    onJsonChange(filtrados)
    return filtrados
  }

  // ── Abrir selector de carpeta (muestra msg inmediato antes del diálogo OS) ──
  function onClickCargarAuto() {
    carpetaOkRef.current = false
    setEsperandoCarpeta(true)
    inputAutoRef.current.click()
    const onFoco = () => {
      window.removeEventListener('focus', onFoco)
      setEsperandoCarpeta(false)
      setEscaneando(true)  // muestra overlay "Leyendo carpeta..." mientras el browser escanea
      setTimeout(() => {
        if (!carpetaOkRef.current) setEscaneando(false)  // usuario canceló
      }, 120000)
    }
    window.addEventListener('focus', onFoco)
  }

  // ── Carpeta automática ────────────────────────────────────────────────────
  function onCarpetaAuto(e) {
    carpetaOkRef.current = true
    setEsperandoCarpeta(false)
    setEscaneando(false)
    setMostrarResumen(false)
    onJsonChange([])
    setBtnCargando(true)
    setClasActivo(true)

    setTimeout(() => {
      const todos = Array.from(e.target.files)
        .filter(f => f.name.endsWith('.json') && f.name.startsWith(RIPS_PREFIJO))
        .sort((a, b) => b.lastModified - a.lastModified)

      if (todos.length === 0) {
        setClasActivo(false)
        setBtnCargando(false)
        setMostrarResumen(false)
        onJsonChange([])
        return
      }

      const porDia = {}
      todos.forEach(f => {
        const d = fmtFecha(f.lastModified)
        porDia[d] = (porDia[d] || 0) + 1
      })
      const dias = Object.keys(porDia).sort((a, b) => b.localeCompare(a))
      setAutoTodos(todos)
      setDiasPorFecha(dias.map(d => ({ fecha: d, n: porDia[d] })))
      setFechaDesde(dias[dias.length - 1])
      setFechaHasta(dias[0])

      const filtrados = aplicarFiltro(todos, dias[dias.length - 1], dias[0], 1, filtro)
      onJsonChange(filtrados)
      setMostrarResumen(true)
      setClasActivo(false)
      setBtnCargando(false)
    }, 80)
  }

  // ── Cambios de filtro ────────────────────────────────────────────────────
  function onFechaDesde(v) { setFechaDesde(v); aplicarFiltro(autoTodos, v, fechaHasta, numDesde, 'fecha') }
  function onFechaHasta(v) { setFechaHasta(v); aplicarFiltro(autoTodos, fechaDesde, v, numDesde, 'fecha') }
  function onNumDesde(v)   { setNumDesde(v);   aplicarFiltro(autoTodos, fechaDesde, fechaHasta, v, 'numero') }
  function onCambiarFiltro(tipo) {
    setFiltro(tipo)
    aplicarFiltro(autoTodos, fechaDesde, fechaHasta, numDesde, tipo)
  }

  // ── Modo manual ───────────────────────────────────────────────────────────
  function onManual(e) {
    const sel = Array.from(e.target.files).slice(0, RIPS_MAX)
    onJsonChange(sel)
  }

  // ── Excel ─────────────────────────────────────────────────────────────────
  function onExcel(e) {
    onExcelChange(Array.from(e.target.files))
  }

  function cambiarModo(m) {
    setModo(m)
    setMostrarResumen(false)
    onJsonChange([])
    setAutoTodos([])
  }

  const nJson  = jsonFiles.length
  const warn   = nJson >= RIPS_MAX ? `Límite de ${RIPS_MAX} archivos alcanzado.` : null

  return (
    <section className="card upload-card">
      <div className="card-label">
        <svg viewBox="0 0 20 20" fill="currentColor">
          <path d="M3 17a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zM6.293 6.707a1 1 0 010-1.414l3-3a1 1 0 011.414 0l3 3a1 1 0 01-1.414 1.414L11 5.414V13a1 1 0 11-2 0V5.414L7.707 6.707a1 1 0 01-1.414 0z"/>
        </svg>
        Carga de archivos
      </div>

      {/* Overlay procesando archivos */}
      {procesando && (
        <div className="clas-overlay clas-active">
          <div className="clas-card" style={{ minWidth: '320px' }}>
            <div className="clas-spinner">
              <div className="clas-ring"></div>
              <div className="clas-ring clas-ring-2"></div>
              <svg className="clas-icon" viewBox="0 0 24 24" fill="currentColor">
                <path d="M19 3H5C3.9 3 3 3.9 3 5v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10h-4v4h-2v-4H7v-2h4V7h2v4h4v2z"/>
              </svg>
            </div>
            <div className="clas-title">
              {loadingPct < 100 ? 'Enviando archivos al servidor' : 'Analizando con IA'}
            </div>
            <div className="clas-sub">
              {loadingPct < 100
                ? <>{nJson} archivos RIPS · {loadingPct}% subido</>
                : <>Ejecutando malla validadora 2275<span className="clas-dots"><span>.</span><span>.</span><span>.</span></span></>
              }
            </div>
            {/* Barra de progreso */}
            <div style={{ width: '100%', marginTop: '1rem' }}>
              <div style={{
                background: 'rgba(255,255,255,0.15)',
                borderRadius: '999px',
                height: '8px',
                overflow: 'hidden',
              }}>
                <div style={{
                  height: '100%',
                  borderRadius: '999px',
                  background: loadingPct < 100
                    ? 'linear-gradient(90deg, #38bdf8, #818cf8)'
                    : 'linear-gradient(90deg, #34d399, #10b981)',
                  width: loadingPct < 100 ? `${Math.max(loadingPct, 3)}%` : '100%',
                  transition: 'width 0.4s ease',
                  animation: loadingPct >= 100 ? 'pct-pulse 1.5s ease-in-out infinite' : 'none',
                }}/>
              </div>
              <div style={{
                textAlign: 'right',
                fontSize: '0.75rem',
                color: 'rgba(255,255,255,0.7)',
                marginTop: '0.25rem',
              }}>
                {loadingPct < 100 ? `${loadingPct}%` : 'Procesando...'}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Overlay leyendo carpeta (browser escaneando archivos) */}
      {escaneando && (
        <div className="clas-overlay clas-active">
          <div className="clas-card">
            <div className="clas-spinner">
              <div className="clas-ring"></div>
              <div className="clas-ring clas-ring-2"></div>
              <svg className="clas-icon" viewBox="0 0 24 24" fill="currentColor">
                <path d="M22 19a2 2 0 01-2 2H4a2 2 0 01-2-2V5a2 2 0 012-2h5l2 3h9a2 2 0 012 2z"/>
              </svg>
            </div>
            <div className="clas-title">Leyendo carpeta</div>
            <div className="clas-sub">Cargando archivos RIPS<span className="clas-dots"><span>.</span><span>.</span><span>.</span></span></div>
          </div>
        </div>
      )}

      {/* Overlay de clasificación */}
      {clasActivo && (
        <div className="clas-overlay clas-active">
          <div className="clas-card">
            <div className="clas-spinner">
              <div className="clas-ring"></div>
              <div className="clas-ring clas-ring-2"></div>
              <svg className="clas-icon" viewBox="0 0 24 24" fill="currentColor">
                <path d="M19 3H5C3.9 3 3 3.9 3 5v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10h-4v4h-2v-4H7v-2h4V7h2v4h4v2z"/>
              </svg>
            </div>
            <div className="clas-title">Espera un momento por favor</div>
            <div className="clas-sub">Estamos procesando la información<span className="clas-dots"><span>.</span><span>.</span><span>.</span></span></div>
          </div>
        </div>
      )}

      <div className="upload-grid">

        {/* ── Zona JSON ── */}
        <div className={`upload-zone ${nJson > 0 ? 'zone-loaded' : ''}`}>
          <div className="upload-zone-icon">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
              <path d="M22 19a2 2 0 01-2 2H4a2 2 0 01-2-2V5a2 2 0 012-2h5l2 3h9a2 2 0 012 2z"/>
            </svg>
          </div>
          <p className="upload-zone-title">RIPS JSON</p>

          <div className="rips-mode-toggle">
            <button type="button" onClick={() => cambiarModo('auto')}
              className={`rips-mode-btn ${modo === 'auto' ? 'rips-mode-active' : ''}`}>
              <svg viewBox="0 0 16 16" fill="currentColor"><path d="M8 0a8 8 0 100 16A8 8 0 008 0zm3.5 7.5a.5.5 0 010 1H5.707l2.147 2.146a.5.5 0 01-.708.708l-3-3a.5.5 0 010-.708l3-3a.5.5 0 11.708.708L5.707 7.5H11.5z"/></svg>
              Automático
            </button>
            <button type="button" onClick={() => cambiarModo('manual')}
              className={`rips-mode-btn ${modo === 'manual' ? 'rips-mode-active' : ''}`}>
              <svg viewBox="0 0 16 16" fill="currentColor"><path d="M.5 9.9a.5.5 0 01.5.5v2.5a1 1 0 001 1h12a1 1 0 001-1v-2.5a.5.5 0 011 0v2.5a2 2 0 01-2 2H2a2 2 0 01-2-2v-2.5a.5.5 0 01.5-.5z"/><path d="M7.646 1.146a.5.5 0 01.708 0l3 3a.5.5 0 01-.708.708L8.5 2.707V11.5a.5.5 0 01-1 0V2.707L5.354 4.854a.5.5 0 11-.708-.708l3-3z"/></svg>
              Manual
            </button>
          </div>

          {/* Modo automático */}
          {modo === 'auto' && (
            <div>
              <p className="upload-zone-sub">Seleccione la carpeta con archivos <code>Rips_SL*.json</code></p>
              <button type="button"
                className={`btn-file ${(btnCargando || esperandoCarpeta) ? 'btn-cargando' : ''}`}
                onClick={onClickCargarAuto}
                disabled={esperandoCarpeta || btnCargando}>
                Cargue RIPS JSON
                {(btnCargando || esperandoCarpeta) && <span className="btn-cargando-spin"></span>}
              </button>
              <input ref={inputAutoRef} type="file" webkitdirectory="true" mozdirectory="true"
                style={{ display: 'none' }} onChange={onCarpetaAuto} />

              {esperandoCarpeta && (
                <div className="uc-procesando-msg">
                  <span className="uc-procesando-dot"></span>
                  Estamos en procesamiento de la información
                  <span className="clas-dots"><span>.</span><span>.</span><span>.</span></span>
                </div>
              )}

              {mostrarResumen && (
                <div id="autoResumen">
                  <div className="rips-resumen-header">
                    <span>{autoTodos.length} archivos encontrados</span>
                    <span className="rips-resumen-hint">Elija cómo filtrar:</span>
                  </div>
                  <div className="rips-filtro-tabs">
                    <button type="button" onClick={() => onCambiarFiltro('fecha')}
                      className={`rips-ftab ${filtro === 'fecha' ? 'rips-ftab-active' : ''}`}>Por fecha</button>
                    <button type="button" onClick={() => onCambiarFiltro('numero')}
                      className={`rips-ftab ${filtro === 'numero' ? 'rips-ftab-active' : ''}`}>Por N° de archivo</button>
                  </div>

                  {filtro === 'fecha' && (
                    <div className="rips-date-row">
                      <label>Desde</label>
                      <input type="date" value={fechaDesde} onChange={e => onFechaDesde(e.target.value)} />
                      <label>Hasta</label>
                      <input type="date" value={fechaHasta} onChange={e => onFechaHasta(e.target.value)} />
                    </div>
                  )}

                  {filtro === 'numero' && (
                    <div className="rips-num-row">
                      <label>Desde el archivo N°</label>
                      <input type="number" min="1" value={numDesde}
                        onChange={e => onNumDesde(Number(e.target.value))} />
                    </div>
                  )}

                  <div className="rips-dias-wrap">
                    <table className="rips-dias-tbl">
                      <thead><tr><th>Fecha</th><th>Archivos</th></tr></thead>
                      <tbody>
                        {diasPorFecha.map(({ fecha, n }) => (
                          <tr key={fecha}><td>{fecha}</td><td>{n}</td></tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Modo manual */}
          {modo === 'manual' && (
            <div>
              <p className="upload-zone-sub">Seleccione los archivos <code>Rips_SL*.json</code> (máx. {RIPS_MAX})</p>
              <button type="button" className="btn-file"
                onClick={() => inputManualRef.current.click()}>
                Seleccionar archivos
              </button>
              <input ref={inputManualRef} type="file" accept=".json,application/json" multiple
                style={{ display: 'none' }} onChange={onManual} />
              {nJson > 0 && (
                <div className="rips-manual-info" style={{ display: 'block' }}>
                  <span>{nJson} archivo(s) seleccionado(s)</span>
                </div>
              )}
            </div>
          )}

          {/* Contador */}
          {nJson > 0 && (
            <div className="file-counter" style={{ display: 'flex' }}>
              <span className="counter-num">{nJson}</span>
              <span className="counter-lbl">archivo(s) listos</span>
            </div>
          )}
          {warn && (
            <div className="file-counter-warn" style={{ display: 'block' }}>{warn}</div>
          )}
        </div>

        {/* ── Zona Excel ── */}
        <div className="upload-zone">
          <div className="upload-zone-icon excel-icon">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
              <rect x="3" y="3" width="18" height="18" rx="2"/>
              <path d="M3 9h18M9 21V9"/>
            </svg>
          </div>
          <p className="upload-zone-title">Base EPS <span className="tag-optional">Opcional</span></p>
          <p className="upload-zone-sub">Uno o varios Excel de autorizaciones</p>
          <button type="button" className="btn-file btn-excel"
            onClick={() => inputExcelRef.current.click()}>
            Seleccionar Excel
          </button>
          <input ref={inputExcelRef} type="file" accept=".xlsx,.xls" multiple
            style={{ display: 'none' }} onChange={onExcel} />
          {excelFiles.length > 0 && (
            <ul className="file-list">
              {excelFiles.map((f, i) => <li key={i}>{f.name}</li>)}
            </ul>
          )}
        </div>

      </div>{/* /upload-grid */}

      {/* Aviso */}
      {hasResults ? (
        <div className="notice-jorge notice-jorge-warn">
          <svg viewBox="0 0 20 20" fill="currentColor">
            <path fillRule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clipRule="evenodd"/>
          </svg>
          <div>
            <strong>Resultados pendientes de exportar.</strong>
            {' '}Descarga el Excel <em>antes</em> de procesar nuevos archivos o perderás los resultados actuales.
          </div>
        </div>
      ) : (
        <div className="notice-jorge">
          <svg viewBox="0 0 20 20" fill="currentColor">
            <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clipRule="evenodd"/>
          </svg>
          Exporte el Excel antes de procesar un nuevo lote para no perder los resultados.
        </div>
      )}

      {/* Botones */}
      <div className="action-bar">
        <button type="button" className="btn btn-primary"
          onClick={() => { setProcesando(true); onProcesar() }}
          disabled={!nJson || loading}>
          <svg viewBox="0 0 20 20" fill="currentColor">
            <path d="M10 12a2 2 0 100-4 2 2 0 000 4z"/>
            <path fillRule="evenodd" d="M.458 10C1.732 5.943 5.522 3 10 3s8.268 2.943 9.542 7c-1.274 4.057-5.064 7-9.542 7S1.732 14.057.458 10zM14 10a4 4 0 11-8 0 4 4 0 018 0z" clipRule="evenodd"/>
          </svg>
          Procesar archivos
        </button>
        <button type="button" className="btn btn-success" onClick={onExportar}
          disabled={!hasResults}>
          <svg viewBox="0 0 20 20" fill="currentColor">
            <path fillRule="evenodd" d="M3 17a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zm3.293-7.707a1 1 0 011.414 0L9 10.586V3a1 1 0 112 0v7.586l1.293-1.293a1 1 0 111.414 1.414l-3 3a1 1 0 01-1.414 0l-3-3a1 1 0 010-1.414z" clipRule="evenodd"/>
          </svg>
          Exportar Excel
        </button>
      </div>

      {/* Error */}
      {error && (
        <div className="alert alert-error">
          <svg viewBox="0 0 20 20" fill="currentColor">
            <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd"/>
          </svg>
          <span>{error}</span>
        </div>
      )}
    </section>
  )
}
