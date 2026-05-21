import { useState, useEffect } from 'react'
import {
  listarUsuarios, crearUsuario, eliminarUsuario,
  cambiarPasswordUsuario, cambiarMiPassword,
} from '../api/client'

const ROLES = { admin: 'Administrador', auditor: 'Auditor Clínico' }

function IcoUser() {
  return (
    <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '1rem', height: '1rem' }}>
      <path fillRule="evenodd" d="M10 9a3 3 0 100-6 3 3 0 000 6zm-7 9a7 7 0 1114 0H3z" clipRule="evenodd"/>
    </svg>
  )
}
function IcoKey() {
  return (
    <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '1rem', height: '1rem' }}>
      <path fillRule="evenodd" d="M18 8a6 6 0 01-7.743 5.743L10 14l-1 1-1 1H6v2H2v-4l4.257-4.257A6 6 0 1118 8zm-6-4a1 1 0 100 2 2 2 0 012 2 1 1 0 102 0 4 4 0 00-4-4z" clipRule="evenodd"/>
    </svg>
  )
}
function IcoTrash() {
  return (
    <svg viewBox="0 0 20 20" fill="currentColor" style={{ width: '0.9rem', height: '0.9rem' }}>
      <path fillRule="evenodd" d="M9 2a1 1 0 00-.894.553L7.382 4H4a1 1 0 000 2v10a2 2 0 002 2h8a2 2 0 002-2V6a1 1 0 100-2h-3.382l-.724-1.447A1 1 0 0011 2H9zM7 8a1 1 0 012 0v6a1 1 0 11-2 0V8zm5-1a1 1 0 00-1 1v6a1 1 0 102 0V8a1 1 0 00-1-1z" clipRule="evenodd"/>
    </svg>
  )
}

// ── Sección: Gestión de usuarios ─────────────────────────────────────────────

function GestionUsuarios() {
  const [usuarios, setUsuarios]       = useState([])
  const [cargando, setCargando]       = useState(true)
  const [error, setError]             = useState(null)
  const [exito, setExito]             = useState(null)

  // Nuevo usuario
  const [form, setForm] = useState({ username: '', password: '', nombre: '', role: 'auditor' })
  const [creando, setCreando]         = useState(false)

  // Cambio de contraseña inline
  const [editPwd, setEditPwd]         = useState(null)  // username en edición
  const [nuevaPwd, setNuevaPwd]       = useState('')
  const [guardandoPwd, setGuardandoPwd] = useState(false)

  useEffect(() => { cargar() }, [])

  async function cargar() {
    setCargando(true); setError(null)
    try { setUsuarios(await listarUsuarios()) }
    catch (e) { setError(e?.response?.data?.detail ?? 'Error cargando usuarios') }
    finally { setCargando(false) }
  }

  function flash(msg) {
    setExito(msg)
    setTimeout(() => setExito(null), 3500)
  }

  async function handleCrear(e) {
    e.preventDefault()
    setCreando(true); setError(null)
    try {
      await crearUsuario(form)
      setForm({ username: '', password: '', nombre: '', role: 'auditor' })
      await cargar()
      flash(`Usuario "${form.username}" creado.`)
    } catch (e) { setError(e?.response?.data?.detail ?? 'Error creando usuario') }
    finally { setCreando(false) }
  }

  async function handleEliminar(username) {
    if (!window.confirm(`¿Eliminar el usuario "${username}"? Esta acción no se puede deshacer.`)) return
    setError(null)
    try { await eliminarUsuario(username); await cargar(); flash(`Usuario "${username}" eliminado.`) }
    catch (e) { setError(e?.response?.data?.detail ?? 'Error eliminando usuario') }
  }

  async function handleGuardarPwd(username) {
    if (!nuevaPwd) return
    setGuardandoPwd(true); setError(null)
    try {
      await cambiarPasswordUsuario(username, nuevaPwd)
      setEditPwd(null); setNuevaPwd('')
      flash(`Contraseña de "${username}" actualizada.`)
    } catch (e) { setError(e?.response?.data?.detail ?? 'Error cambiando contraseña') }
    finally { setGuardandoPwd(false) }
  }

  return (
    <>
      {/* ── Tabla de usuarios ── */}
      <div className="admin-section">
        <h3 className="admin-section-title">Usuarios registrados</h3>
        {cargando ? (
          <p className="admin-hint">Cargando...</p>
        ) : (
          <div className="admin-table-wrap">
            <table className="admin-table">
              <thead>
                <tr><th>Usuario</th><th>Nombre</th><th>Rol</th><th>Acciones</th></tr>
              </thead>
              <tbody>
                {usuarios.map(u => (
                  <tr key={u.username}>
                    <td><code>{u.username}</code></td>
                    <td>{u.nombre || '—'}</td>
                    <td>
                      <span className={`admin-role-tag admin-role-${u.role}`}>
                        {ROLES[u.role] ?? u.role}
                      </span>
                    </td>
                    <td className="admin-actions">
                      {editPwd === u.username ? (
                        <div className="admin-pwd-inline">
                          <input
                            type="password"
                            placeholder="Nueva contraseña (mín. 8)"
                            value={nuevaPwd}
                            onChange={e => setNuevaPwd(e.target.value)}
                            className="admin-input-sm"
                            minLength={8}
                            autoFocus
                          />
                          <button className="btn-admin-ok" onClick={() => handleGuardarPwd(u.username)}
                            disabled={guardandoPwd || nuevaPwd.length < 8}>
                            {guardandoPwd ? '...' : 'Guardar'}
                          </button>
                          <button className="btn-admin-cancel" onClick={() => { setEditPwd(null); setNuevaPwd('') }}>
                            Cancelar
                          </button>
                        </div>
                      ) : (
                        <>
                          <button className="btn-admin-icon btn-admin-key" title="Cambiar contraseña"
                            onClick={() => { setEditPwd(u.username); setNuevaPwd('') }}>
                            <IcoKey /> Contraseña
                          </button>
                          <button className="btn-admin-icon btn-admin-del" title="Eliminar usuario"
                            onClick={() => handleEliminar(u.username)}>
                            <IcoTrash /> Eliminar
                          </button>
                        </>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* ── Crear usuario ── */}
      <div className="admin-section">
        <h3 className="admin-section-title">Crear nuevo usuario</h3>
        <form className="admin-form" onSubmit={handleCrear}>
          <div className="admin-form-grid">
            <label className="admin-label">
              Usuario
              <input className="admin-input" type="text" required value={form.username}
                onChange={e => setForm(f => ({ ...f, username: e.target.value }))}
                placeholder="nombre_usuario" autoComplete="off"/>
            </label>
            <label className="admin-label">
              Nombre completo
              <input className="admin-input" type="text" value={form.nombre}
                onChange={e => setForm(f => ({ ...f, nombre: e.target.value }))}
                placeholder="Nombre Apellido"/>
            </label>
            <label className="admin-label">
              Contraseña
              <input className="admin-input" type="password" required minLength={8} value={form.password}
                onChange={e => setForm(f => ({ ...f, password: e.target.value }))}
                placeholder="Mínimo 8 caracteres" autoComplete="new-password"/>
            </label>
            <label className="admin-label">
              Rol
              <select className="admin-input" value={form.role}
                onChange={e => setForm(f => ({ ...f, role: e.target.value }))}>
                <option value="auditor">Auditor Clínico</option>
                <option value="admin">Administrador</option>
              </select>
            </label>
          </div>
          <button className="btn btn-primary" type="submit" disabled={creando}>
            {creando ? 'Creando...' : 'Crear usuario'}
          </button>
        </form>
      </div>

      {/* Mensajes */}
      {error  && <div className="alert alert-error"  style={{ marginTop: '1rem' }}><span>{error}</span></div>}
      {exito  && <div className="alert alert-ok"     style={{ marginTop: '1rem' }}><span>{exito}</span></div>}
    </>
  )
}

// ── Sección: Cambiar mi contraseña ───────────────────────────────────────────

function MiContrasena() {
  const [actual,   setActual]   = useState('')
  const [nueva,    setNueva]    = useState('')
  const [confirma, setConfirma] = useState('')
  const [enviando, setEnviando] = useState(false)
  const [msg,      setMsg]      = useState(null)  // { type, text }

  async function handleSubmit(e) {
    e.preventDefault()
    setMsg(null)
    if (nueva !== confirma) { setMsg({ type: 'error', text: 'Las contraseñas nuevas no coinciden.' }); return }
    if (nueva.length < 8)   { setMsg({ type: 'error', text: 'La contraseña debe tener al menos 8 caracteres.' }); return }
    setEnviando(true)
    try {
      await cambiarMiPassword(actual, nueva)
      setActual(''); setNueva(''); setConfirma('')
      setMsg({ type: 'ok', text: 'Contraseña actualizada correctamente.' })
    } catch (e) {
      setMsg({ type: 'error', text: e?.response?.data?.detail ?? 'Error cambiando contraseña.' })
    } finally { setEnviando(false) }
  }

  return (
    <div className="admin-section admin-section-narrow">
      <h3 className="admin-section-title">Cambiar mi contraseña</h3>
      <form className="admin-form" onSubmit={handleSubmit}>
        <label className="admin-label">
          Contraseña actual
          <input className="admin-input" type="password" required value={actual}
            onChange={e => setActual(e.target.value)} autoComplete="current-password"/>
        </label>
        <label className="admin-label">
          Nueva contraseña
          <input className="admin-input" type="password" required minLength={8} value={nueva}
            onChange={e => setNueva(e.target.value)}
            placeholder="Mínimo 8 caracteres" autoComplete="new-password"/>
        </label>
        <label className="admin-label">
          Confirmar nueva contraseña
          <input className="admin-input" type="password" required minLength={8} value={confirma}
            onChange={e => setConfirma(e.target.value)} autoComplete="new-password"/>
        </label>
        <button className="btn btn-primary" type="submit" disabled={enviando}>
          {enviando ? 'Guardando...' : 'Actualizar contraseña'}
        </button>
      </form>
      {msg && (
        <div className={`alert ${msg.type === 'ok' ? 'alert-ok' : 'alert-error'}`} style={{ marginTop: '1rem' }}>
          <span>{msg.text}</span>
        </div>
      )}
    </div>
  )
}

// ── Componente principal Admin ────────────────────────────────────────────────

export default function Admin({ usuario }) {
  const [tab, setTab] = useState(usuario?.role === 'admin' ? 'usuarios' : 'password')
  const esAdmin = usuario?.role === 'admin'

  return (
    <section className="card admin-card">
      <div className="card-label">
        <svg viewBox="0 0 20 20" fill="currentColor">
          <path fillRule="evenodd" d="M11.49 3.17c-.38-1.56-2.6-1.56-2.98 0a1.532 1.532 0 01-2.286.948c-1.372-.836-2.942.734-2.106 2.106.54.886.061 2.042-.947 2.287-1.561.379-1.561 2.6 0 2.978a1.532 1.532 0 01.947 2.287c-.836 1.372.734 2.942 2.106 2.106a1.532 1.532 0 012.287.947c.379 1.561 2.6 1.561 2.978 0a1.533 1.533 0 012.287-.947c1.372.836 2.942-.734 2.106-2.106a1.533 1.533 0 01.947-2.287c1.561-.379 1.561-2.6 0-2.978a1.532 1.532 0 01-.947-2.287c.836-1.372-.734-2.942-2.106-2.106a1.532 1.532 0 01-2.287-.947zM10 13a3 3 0 100-6 3 3 0 000 6z" clipRule="evenodd"/>
        </svg>
        Administración
      </div>

      <div className="admin-tabs">
        {esAdmin && (
          <button type="button" className={`admin-tab${tab === 'usuarios' ? ' admin-tab-active' : ''}`}
            onClick={() => setTab('usuarios')}>
            <IcoUser /> Usuarios
          </button>
        )}
        <button type="button" className={`admin-tab${tab === 'password' ? ' admin-tab-active' : ''}`}
          onClick={() => setTab('password')}>
          <IcoKey /> Mi contraseña
        </button>
      </div>

      <div className="admin-body">
        {tab === 'usuarios' && esAdmin && <GestionUsuarios />}
        {tab === 'password' && <MiContrasena />}
      </div>
    </section>
  )
}
