// ResultsSection — replica exacta del dashboard + tablas del template Flask
// Usa las mismas clases CSS de static/styles.css

function pct(n, t) { return t > 0 ? Math.round((n / t) * 100) : 0 }

function TopReglas({ reglas }) {
  if (!reglas || !reglas.length) return null
  const max = reglas[0][1]
  return (
    <div className="mc-top-rules">
      <div className="mtr-header">Reglas más frecuentes</div>
      {reglas.map(([regla, cnt]) => (
        <div key={regla} className="mtr-row">
          <span className="mtr-code">{regla}</span>
          <div className="mtr-bar-wrap">
            <div className="mtr-bar" style={{ width: `${Math.round((cnt / max) * 100)}%` }}></div>
          </div>
          <span className="mtr-cnt">{cnt}</span>
        </div>
      ))}
    </div>
  )
}

function ModuleCard({ title, sub, iconClass = '', total, criticas, notificaciones, topReglas, colorClass }) {
  const cls = colorClass || (criticas > 0 ? 'mc-danger' : total > 0 ? 'mc-warn' : 'mc-ok')
  const cntCls = criticas > 0 ? 'mc-count-red' : total > 0 ? 'mc-count-warn' : 'mc-count-ok'
  const t = total || 1
  return (
    <div className={`module-card ${cls}`}>
      <div className="mc-header">
        <div className={`mc-icon ${iconClass}`}></div>
        <div>
          <div className="mc-title">{title}</div>
          <div className="mc-sub">{sub}</div>
        </div>
      </div>
      <div className="mc-count-row">
        <span className={`mc-count ${cntCls}`}>{total}</span>
        <span className="mc-count-lbl">alertas</span>
      </div>
      {total > 0 ? (
        <>
          <div className="mc-sev-bars">
            <div className="sev-row">
              <span className="sev-lbl sev-red">Crítica</span>
              <div className="sev-track"><div className="sev-fill sev-fill-red" style={{ width: `${pct(criticas, t)}%` }}></div></div>
              <span className="sev-num">{criticas}</span>
            </div>
            <div className="sev-row">
              <span className="sev-lbl sev-amber">Alta/Med</span>
              <div className="sev-track"><div className="sev-fill sev-fill-amber" style={{ width: `${pct(notificaciones, t)}%` }}></div></div>
              <span className="sev-num">{notificaciones}</span>
            </div>
          </div>
          <TopReglas reglas={topReglas} />
        </>
      ) : (
        <div className="mc-ok-msg">
          <svg viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd"/></svg>
          Sin alertas detectadas
        </div>
      )}
    </div>
  )
}

function SevBadge({ v }) {
  if (v === 'critica') return <span className="badge-sev sev-critica">Crítica</span>
  if (v === 'alta')    return <span className="badge-sev sev-alta">Alta</span>
  return <span className="badge-sev sev-media">Media</span>
}

function RowCls(v) {
  if (v === 'critica') return 'row-critica'
  if (v === 'alta')    return 'row-alta'
  if (v === 'media')   return 'row-media'
  return ''
}

// ── Tabla genérica de alertas ─────────────────────────────────────────────
function AlertTable({ title, desc, items = [], tipo, labelCls = 'label-danger', cardCls = 'alert-card-danger', ico }) {
  if (!items || items.length === 0) return null
  const badgeCls = labelCls.replace('label-', 'badge-')

  const ICO_DANGER = (
    <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd"/>
  )
  const ICO_WARN = (
    <path fillRule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clipRule="evenodd"/>
  )
  const ICO_INFO = (
    <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clipRule="evenodd"/>
  )
  const icoPath = ico === 'warn' ? ICO_WARN : ico === 'info' ? ICO_INFO : ICO_DANGER

  return (
    <section className={`card alert-card ${cardCls}`}>
      <div className={`card-label ${labelCls}`}>
        <svg viewBox="0 0 20 20" fill="currentColor">{icoPath}</svg>
        {title}
        <span className={`badge-count ${badgeCls}`}>{items.length}</span>
      </div>
      {desc && <p className="alert-desc" dangerouslySetInnerHTML={{ __html: desc }}></p>}
      <div className="table-wrapper">
        <table>
          <thead><tr>{COLS[tipo].map(c => <th key={c}>{c}</th>)}</tr></thead>
          <tbody>
            {items.map((item, i) => (
              <tr key={i}>
                <td className="text-muted">{i + 1}</td>
                {ROWS[tipo](item)}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  )
}

// ── Definición de columnas y celdas por tipo de tabla ─────────────────────
const COLS = {
  tipo_doc:        ['#','N° Documento','Tipo RIPS','Tipo EPS','Autorización','Archivo EPS'],
  amb_par:         ['#','Mensaje','Código Proc.','N° Autorización RIPS','N° Documento','Factura RIPS','Archivo RIPS','Archivo EPS'],
  hosp_no_rel:     ['#','Mensaje','N° Autorización EPS','N° Documento','Fecha Emisión EPS','Ingreso','Egreso','Factura RIPS','Archivo EPS'],
  estancia_sin:    ['#','Mensaje','N° Auth Estancia (RIPS)','N° Documento','Factura RIPS','Archivo RIPS','Archivo EPS'],
  proc_qx:         ['#','Mensaje','Código Proc.','N° Autorización','N° Documento','Factura RIPS','Archivo RIPS','Archivo EPS'],
  sin_aut_rel:     ['#','Mensaje','N° Documento','Factura RIPS','Archivo RIPS','Archivo EPS'],
  amb_emision:     ['#','N° Autorización','N° Documento','Fecha Servicio RIPS','Fecha Emisión EPS','Archivo RIPS','Archivo EPS'],
  hosp_sin_aut:    ['#','Mensaje','N° Documento','Factura RIPS','Archivo RIPS','Archivo EPS'],
  hosp_no_cruza:   ['#','Código','N° Autorización','N° Documento','Sección','Factura RIPS','Archivo RIPS','Archivo EPS'],
  hosp_proc_nc:    ['#','Mensaje','Código Proc.','N° Autorización RIPS','N° Documento','Factura RIPS','Archivo RIPS','Archivo EPS'],
  hosp_cups_dup:   ['#','Mensaje','Código CUPS','Repeticiones','N° Documento','Factura RIPS','Archivo RIPS'],
  proc_sin_aut:    ['#','Código Proc.','N° Documento','Fecha Servicio','Sección','Factura RIPS','Archivo RIPS'],
  proc_aut_nc:     ['#','Código CUPS','N° Autorización RIPS','N° Documento','Fecha Servicio','Factura RIPS','Archivo RIPS','Archivo EPS'],
  cups_ne:         ['#','Código Proc.','N° Documento','Fecha Servicio','Sección','Factura RIPS','Archivo RIPS'],
  auditoria:       ['#','ID Regla','Severidad','N° Documento','Factura','Campo','Mensaje','Valor Actual','Archivo'],
  medicamentos:    ['#','Archivo','N° Factura','ID RIPS','Código','Nombre Tecnología','Valor Servicio'],
}

function M(v, cls = '') { return <td className={`mono ${cls}`}>{v}</td> }
function Sm(v)          { return <td className="text-muted small">{v}</td> }
function T(v)           { return <td>{v}</td> }

const ROWS = {
  tipo_doc:    i => (<>{M(i.num_doc)}<td><span className="tag tag-rips">{i.tipo_rips}</span></td><td><span className="tag tag-eps">{i.tipo_eps}</span></td>{M(i.num_aut)}{Sm(i.archivo_eps)}</>),
  amb_par:     i => (<>{T(i.mensaje)}{M(i.cod_proc)}{M(i.num_aut)}{M(i.num_doc)}{M(i.num_factura)}{Sm(i.archivo_rips)}{Sm(i.archivo_eps)}</>),
  hosp_no_rel: i => (<>{T(i.mensaje)}{M(i.num_aut)}{M(i.num_doc)}{M(i.fecha_emision)}{M(i.fecha_ingreso)}{M(i.fecha_egreso)}{M(i.num_factura)}{Sm(i.archivo_eps)}</>),
  estancia_sin:i => (<>{T(i.mensaje)}{M(i.auths_hosp)}{M(i.num_doc)}{M(i.num_factura)}{Sm(i.archivo_rips)}{Sm(i.archivo_eps)}</>),
  proc_qx:     i => (<>{T(i.mensaje)}{M(i.cod_proc)}{M(i.num_aut)}{M(i.num_doc)}{M(i.num_factura)}{Sm(i.archivo_rips)}{Sm(i.archivo_eps)}</>),
  sin_aut_rel: i => (<>{T(i.mensaje)}{M(i.num_doc)}{M(i.num_factura)}{Sm(i.archivo_rips)}{Sm(i.archivo_eps)}</>),
  amb_emision: i => (<>{M(i.num_aut)}{M(i.num_doc)}{M(i.fecha_servicio)}{M(i.fecha_emision)}{Sm(i.archivo_rips)}{Sm(i.archivo_eps)}</>),
  hosp_sin_aut:i => (<>{T(i.mensaje)}{M(i.num_doc)}{M(i.num_factura)}{Sm(i.archivo_rips)}{Sm(i.archivo_eps)}</>),
  hosp_no_cruza:i=> (<>{M(i.cod_proc)}{M(i.num_aut)}{M(i.num_doc)}{Sm(i.seccion)}{M(i.num_factura)}{Sm(i.archivo_rips)}{Sm(i.archivo_eps)}</>),
  hosp_proc_nc:i => (<>{T(i.mensaje)}{M(i.cod_proc)}{M(i.num_aut)}{M(i.num_doc)}{M(i.num_factura)}{Sm(i.archivo_rips)}{Sm(i.archivo_eps)}</>),
  hosp_cups_dup:i=> (<>{T(i.mensaje)}{M(i.cod_proc)}<td className="text-center">{i.repeticiones}</td>{M(i.num_doc)}{M(i.num_factura)}{Sm(i.archivo_rips)}</>),
  proc_sin_aut:i => (<>{M(i.cod_proc)}{M(i.num_doc)}{M(i.fecha_inicio)}{Sm(i.seccion)}{M(i.num_factura)}{Sm(i.archivo_rips)}</>),
  proc_aut_nc: i => (<>{M(i.cod_proc)}{M(i.num_aut)}{M(i.num_doc)}{M(i.fecha_inicio)}{M(i.num_factura)}{Sm(i.archivo_rips)}{Sm(i.archivo_eps)}</>),
  cups_ne:     i => (<>{M(i.cod_proc)}{M(i.num_doc)}{M(i.fecha_inicio)}{Sm(i.seccion)}{M(i.num_factura)}{Sm(i.archivo_rips)}</>),
  auditoria:   i => (<><td><span className="tag tag-code">{i.id_regla}</span></td><td><SevBadge v={i.severidad}/></td>{M(i.num_doc)}{M(i.num_factura)}{Sm(i.campo)}{T(i.mensaje)}{M(i.valor_actual)}{Sm(i.archivo)}</>),
  medicamentos:i => (<>{Sm(i.archivo)}{M(i.numeroFactura)}{M(i.idrips)}<td><span className="tag tag-code">{i.codConsulta}</span></td><td className="text-warn">{i.nomTecnologiaSalud}</td><td className="align-right mono">{i.vrServicio !== '' && i.vrServicio !== undefined ? Number(i.vrServicio).toLocaleString('es-CO', {maximumFractionDigits:0}) : ''}</td></>),
}

// ── Componente principal ──────────────────────────────────────────────────
export default function ResultsSection({ resultado }) {
  const { stats = {}, registros = [], alertas = {}, validaciones_malla = [], validaciones_general = [], validaciones_auditoria = [] } = resultado

  const grandTotal   = (stats.malla_total || 0) + (stats.general_total || 0) + (stats.auditoria_total || 0)
  const grandCritica = (stats.malla_criticas || 0) + (stats.general_criticas || 0) + (stats.auditoria_criticas || 0)

  const noAlertasAut = !alertas || Object.values(alertas).every(v => !v || v.length === 0)

  return (
    <>
      {/* ── Dashboard hero ── */}
      <section className="dashboard-section">
        <div className="dashboard-hero">
          <div className="dh-left">
            <div className="dh-title">
              <svg viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M3 3a1 1 0 000 2v8a2 2 0 002 2h2.586l-1.293 1.293a1 1 0 101.414 1.414L10 15.414l2.293 2.293a1 1 0 001.414-1.414L12.414 15H15a2 2 0 002-2V5a1 1 0 100-2H3zm11 4a1 1 0 10-2 0v4a1 1 0 102 0V7zm-3 1a1 1 0 10-2 0v3a1 1 0 102 0V8zM8 9a1 1 0 00-2 0v2a1 1 0 102 0V9z" clipRule="evenodd"/>
              </svg>
              Resultados del análisis
            </div>
            <div className="dh-meta">
              <span className="dh-tag">
                <svg viewBox="0 0 16 16" fill="currentColor"><path d="M14 4.5V14a2 2 0 01-2 2H4a2 2 0 01-2-2V2a2 2 0 012-2h5.5L14 4.5z"/></svg>
                {stats.archivos_json} archivo(s) JSON
              </span>
              <span className="dh-tag">
                <svg viewBox="0 0 16 16" fill="currentColor"><path d="M1.5 1a.5.5 0 00-.5.5v3a.5.5 0 01-1 0v-3A1.5 1.5 0 011.5 0h3a.5.5 0 010 1h-3zm9 0a.5.5 0 000 1h3a.5.5 0 01.5.5v3a.5.5 0 001 0v-3A1.5 1.5 0 0013.5 0h-3zm-9 13a.5.5 0 01.5-.5h3a.5.5 0 010 1h-3a1.5 1.5 0 01-1.5-1.5v-3a.5.5 0 011 0v3zm13 0a1.5 1.5 0 01-1.5 1.5h-3a.5.5 0 010-1h3a.5.5 0 00.5-.5v-3a.5.5 0 011 0v3z"/></svg>
                {stats.total_rips} registros RIPS
              </span>
              <span className="dh-tag dh-tag-time">
                <svg viewBox="0 0 16 16" fill="currentColor"><path d="M8 3.5a.5.5 0 00-1 0V9a.5.5 0 00.252.434l3.5 2a.5.5 0 00.496-.868L8 8.71V3.5z"/><path d="M8 16A8 8 0 108 0a8 8 0 000 16zm7-8A7 7 0 111 8a7 7 0 0114 0z"/></svg>
                {stats.tiempo_procesamiento}
              </span>
            </div>
          </div>
          <div className="dh-totals">
            <div className={`dh-total-big ${grandTotal > 0 ? 'dh-has-alerts' : ''}`}>
              <span className="dh-big-num">{grandTotal}</span>
              <span className="dh-big-lbl">alertas totales</span>
            </div>
            {grandCritica > 0 && (
              <div className="dh-critica-badge">
                <svg viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clipRule="evenodd"/></svg>
                {grandCritica} crítica(s)
              </div>
            )}
          </div>
        </div>

        {/* Module cards */}
        <div className="module-grid">
          <ModuleCard title="Malla RIPS 2275" sub="Validación estructural y normativa"
            total={stats.malla_total} criticas={stats.malla_criticas} notificaciones={stats.malla_notificaciones} topReglas={stats.malla_top_reglas} />
          <ModuleCard title="Validación General" sub="Cruce autorizaciones EPS–RIPS" iconClass="mc-icon-general"
            total={stats.general_total} criticas={stats.general_criticas} notificaciones={stats.general_notificaciones} topReglas={stats.general_top_reglas} />
          <ModuleCard title="Auditoría Clínica" sub="Estancia · Transfusional · Urgencias · Lab" iconClass="mc-icon-audit"
            total={stats.auditoria_total} criticas={stats.auditoria_criticas} notificaciones={stats.auditoria_notificaciones} topReglas={stats.auditoria_top_reglas} />
        </div>
      </section>

      {/* ── KPI chips ── */}
      <section className="stats-row">
        <div className="stat-chip"><span className="stat-val">{stats.archivos_json}</span><span className="stat-lbl">JSON procesados</span></div>
        <div className={`stat-chip ${stats.alerta_volumen ? 'stat-danger' : ''}`}><span className="stat-val">{stats.total_rips}</span><span className="stat-lbl">Registros RIPS</span></div>
        <div className={`stat-chip ${stats.med_invalidos > 0 ? 'stat-warn' : ''}`}><span className="stat-val">{stats.med_invalidos}</span><span className="stat-lbl">Med. inválidos</span></div>
        <div className="stat-chip"><span className="stat-val">{stats.auts_rips}</span><span className="stat-lbl">Autorizaciones RIPS</span></div>
        {stats.auts_excel > 0 && <>
          <div className="stat-chip"><span className="stat-val">{stats.auts_excel}</span><span className="stat-lbl">Autorizaciones EPS</span></div>
          {stats.tipo_mismatch > 0 && <div className="stat-chip stat-danger"><span className="stat-val">{stats.tipo_mismatch}</span><span className="stat-lbl">Tipo doc. difiere</span></div>}
          {stats.amb_par_nc > 0    && <div className="stat-chip stat-danger"><span className="stat-val">{stats.amb_par_nc}</span><span className="stat-lbl">Par auth/cód no cruza (amb)</span></div>}
          {stats.hosp_aut_no_rel > 0 && <div className="stat-chip stat-danger"><span className="stat-val">{stats.hosp_aut_no_rel}</span><span className="stat-lbl">Auth no relacionada (hosp)</span></div>}
          {stats.amb_emision_post > 0 && <div className="stat-chip stat-danger"><span className="stat-val">{stats.amb_emision_post}</span><span className="stat-lbl">Emisión posterior (amb)</span></div>}
          {stats.hosp_cod_sin_aut > 0 && <div className="stat-chip stat-danger"><span className="stat-val">{stats.hosp_cod_sin_aut}</span><span className="stat-lbl">Cód. sin auth (hosp)</span></div>}
          {stats.hosp_proc_no_cruza > 0 && <div className="stat-chip stat-danger"><span className="stat-val">{stats.hosp_proc_no_cruza}</span><span className="stat-lbl">Cód. no cruza (hosp)</span></div>}
          {stats.proc_sin_aut_amb > 0  && <div className="stat-chip stat-danger"><span className="stat-val">{stats.proc_sin_aut_amb}</span><span className="stat-lbl">Proc. sin auth (amb)</span></div>}
          {stats.cups_noestandar_nc > 0 && <div className="stat-chip stat-warn"><span className="stat-val">{stats.cups_noestandar_nc}</span><span className="stat-lbl">CUPS no estándar s/auth</span></div>}
          {stats.hosp_cups_duplicado > 0 && <div className="stat-chip stat-danger"><span className="stat-val">{stats.hosp_cups_duplicado}</span><span className="stat-lbl">CUPS duplicado (hosp)</span></div>}
        </>}
      </section>

      {/* ── Alerta volumen ── */}
      {stats.alerta_volumen && (
        <section className="card alert-card alert-card-vol">
          <div className="card-label label-vol">
            <svg viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clipRule="evenodd"/></svg>
            Volumen RIPS elevado — {stats.total_rips} registros analizados
          </div>
          <p className="alert-desc">Se superaron los <strong>800 registros</strong>. Verifique que el RIPS corresponda a un único periodo de facturación.</p>
        </section>
      )}

      {/* ── Tablas de autorizaciones ── */}
      {alertas && <>
        <AlertTable tipo="tipo_doc"    title="Tipo de documento no coincide con la EPS"                              items={alertas.tipo_doc_mismatch}         ico="warn" labelCls="label-warn" cardCls="alert-card-warn" />
        <AlertTable tipo="amb_par"     title="Ambulatorio / Cirugía ambulatoria: autorización o código no coinciden" items={alertas.amb_par_no_cruza}           ico="danger" desc="El par <code>numAutorizacion</code> + <code>codProcedimiento</code> del RIPS no se encontró en la base de datos EPS." />
        <AlertTable tipo="hosp_no_rel" title="Hospitalario: autorizaciones de la EPS no relacionadas en el RIPS"    items={alertas.hosp_aut_no_relacionada}    ico="info"   labelCls="label-blue" cardCls="alert-card-blue" />
        <AlertTable tipo="estancia_sin" title="Estancia hospitalaria sin autorización en base de datos"               items={alertas.estancia_sin_aut}           ico="danger" desc="El registro de hospitalización no tiene <code>numAutorizacion</code> en la base EPS." />
        <AlertTable tipo="proc_qx"     title="Procedimiento quirúrgico con autorización igual a la de la estancia"   items={alertas.proc_qx_misma_aut_hosp}    ico="warn" labelCls="label-warn" cardCls="alert-card-warn" />
        <AlertTable tipo="sin_aut_rel" title="Sin números de autorización relacionados"                               items={alertas.sin_num_aut_relacionado}    ico="warn" labelCls="label-warn" cardCls="alert-card-warn" />
        <AlertTable tipo="amb_emision" title="Autorización emitida POSTERIOR al servicio (Ambulatorio)"               items={alertas.amb_aut_emision_posterior}  ico="warn" labelCls="label-warn" cardCls="alert-card-warn" />
        <AlertTable tipo="hosp_sin_aut" title="Hospitalización: códigos sin autorización asociada"                    items={alertas.hosp_cod_sin_aut}           ico="danger" />
        <AlertTable tipo="hosp_no_cruza" title="Hospitalización: código no encontrado en base de datos EPS"           items={alertas.hosp_proc_no_cruza}         ico="danger" />
        <AlertTable tipo="hosp_proc_nc" title="Hospitalización + Procedimientos: autorización y código no coinciden"  items={alertas.hosp_proc_cod_no_cruza}    ico="danger" />
        <AlertTable tipo="hosp_cups_dup" title="Hospitalización: código CUPS duplicado"                               items={alertas.hosp_cups_duplicado}        ico="danger" />
        <AlertTable tipo="proc_sin_aut" title="Procedimiento ambulatorio sin autorización obligatoria"                items={alertas.proc_sin_aut_amb}           ico="danger" />
        <AlertTable tipo="proc_aut_nc"  title="Autorización no corresponde al CUPS del RIPS (Ambulatorio)"            items={alertas.proc_aut_no_cruza_amb}      ico="danger" />
        <AlertTable tipo="cups_ne"      title="CUPS no estándar sin autorización asociada"                            items={alertas.cups_noestandar_sin_aut}    ico="warn" labelCls="label-warn" cardCls="alert-card-warn" />

        {noAlertasAut && (
          <section className="card alert-card alert-card-ok">
            <div className="card-label label-ok">
              <svg viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd"/></svg>
              Autorizaciones — Sin inconsistencias detectadas
            </div>
          </section>
        )}
      </>}

      {/* ── Auditoría clínica ── */}
      {validaciones_auditoria.length > 0 && (
        <section className="card results-card">
          <div className="card-label label-danger">
            <svg viewBox="0 0 20 20" fill="currentColor"><path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clipRule="evenodd"/></svg>
            Auditoría Clínica — Hallazgos
            <span className="badge-count badge-danger">{validaciones_auditoria.length}</span>
          </div>
          <div className="table-wrapper">
            <table>
              <thead><tr>{COLS.auditoria.map(c => <th key={c}>{c}</th>)}</tr></thead>
              <tbody>
                {validaciones_auditoria.map((v, i) => (
                  <tr key={i} className={RowCls(v.severidad)}>
                    <td className="text-muted">{i + 1}</td>
                    {ROWS.auditoria(v)}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      {/* ── Medicamentos inválidos ── */}
      {registros.length > 0 && (
        <section className="card results-card">
          <div className="card-label">
            <svg viewBox="0 0 20 20" fill="currentColor"><path d="M9 2a1 1 0 000 2h2a1 1 0 100-2H9z"/><path fillRule="evenodd" d="M4 5a2 2 0 012-2 3 3 0 003 3h2a3 3 0 003-3 2 2 0 012 2v11a2 2 0 01-2 2H6a2 2 0 01-2-2V5zm3 4a1 1 0 000 2h.01a1 1 0 100-2H7zm3 0a1 1 0 000 2h3a1 1 0 100-2h-3zm-3 4a1 1 0 100 2h.01a1 1 0 100-2H7zm3 0a1 1 0 100 2h3a1 1 0 100-2h-3z" clipRule="evenodd"/></svg>
            Medicamentos inválidos encontrados
            <span className="badge-count badge-warn">{registros.length}</span>
          </div>
          <div className="table-wrapper">
            <table>
              <thead><tr>{COLS.medicamentos.map(c => <th key={c}>{c}</th>)}</tr></thead>
              <tbody>
                {registros.map((r, i) => (
                  <tr key={i}>
                    <td className="text-muted">{i + 1}</td>
                    {ROWS.medicamentos(r)}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}
    </>
  )
}
