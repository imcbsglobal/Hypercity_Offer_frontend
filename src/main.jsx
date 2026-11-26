import React, { useState, useEffect, useCallback, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import Cropper from 'react-easy-crop';
import * as XLSX from 'xlsx';
import {
  CalendarDays,
  Bell,
  ClipboardList,
  CheckCircle2,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Clock3,
  Compass,
  Crop,
  Download,
  Edit3,
  ExternalLink,
  Eye,
  EyeOff,
  Filter,
  Gauge,
  Gift,
  Grid2X2Plus,
  Home,
  Image,
  Layers3,
  LineChart,
  List,
  LockKeyhole,
  LogIn,
  LogOut,
  Mail,
  MapPin,
  MoreVertical,
  PackageOpen,
  Phone,
  Plus,
  Search,
  ShieldCheck,
  Sparkles,
  Store,
  Target,
  Trash2,
  Users,
  GripVertical,
  Inbox,
  X,
} from 'lucide-react';
import './styles.css';
import { CustomFormsPage, PublicFormPage } from './customForms';

const API_BASE_URL = (import.meta.env.VITE_API_BASE_URL || '').replace(/\/+$/, '');
const AUTH_STORAGE_KEY = 'admin-auth';
const ACC_MASTER_API_URL = `${(import.meta.env.VITE_ACC_MASTER_API_URL || 'https://hs.hypercityllp.in/api/synctool/acc-master').replace(/\/+$/, '')}/`;
const OFFER_IMAGE_ASPECT = 9 / 16;
const BANNER_IMAGE_ASPECT = 16 / 9;
const CROP_PRESETS = [
  { label: 'Poster 9:16', value: 9 / 16 },
  { label: 'Banner 16:9', value: 16 / 9 },
  { label: 'Square 1:1', value: 1 },
  { label: 'Free', value: null },
];

function readStoredAuth() {
  try {
    const parsed = JSON.parse(localStorage.getItem(AUTH_STORAGE_KEY) || 'null');
    return parsed && parsed.access && parsed.refresh ? parsed : null;
  } catch {
    return null;
  }
}

let session = readStoredAuth();
let refreshInFlight = null;
const sessionListeners = new Set();

function getSession() {
  return session;
}

function setSession(next) {
  session = next;
  try {
    if (next) localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(next));
    else localStorage.removeItem(AUTH_STORAGE_KEY);
  } catch {
    /* storage unavailable, session stays in memory only */
  }
  sessionListeners.forEach((fn) => fn(next));
}

function subscribeToSession(fn) {
  sessionListeners.add(fn);
  return () => { sessionListeners.delete(fn); };
}

function refreshSession() {
  if (refreshInFlight) return refreshInFlight;

  const currentRefresh = session?.refresh;
  if (!currentRefresh) {
    setSession(null);
    return Promise.resolve(null);
  }

  refreshInFlight = fetch(`${API_BASE_URL}/api/auth/refresh/`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ refresh: currentRefresh }),
  })
    .then(async (res) => {
      if (res.ok) return res.json();
      if (res.status === 400 || res.status === 401) setSession(null);
      return null;
    })
    .then((data) => {
      if (!data?.access) return null;
      const next = { ...session, access: data.access, refresh: data.refresh || session.refresh };
      setSession(next);
      return next;
    })
    .catch(() => null)
    .finally(() => { refreshInFlight = null; });

  return refreshInFlight;
}

async function apiFetch(path, options = {}) {
  const headers = { ...(options.headers || {}) };
  const access = session?.access;
  if (access) headers.Authorization = `Bearer ${access}`;
  const url = /^https?:\/\//i.test(path) ? path : `${API_BASE_URL}${path}`;

  const res = await fetch(url, { ...options, headers });

  if (res.status === 401 && session?.refresh) {
    const refreshed = await refreshSession();
    if (refreshed) {
      return fetch(url, {
        ...options,
        headers: { ...headers, Authorization: `Bearer ${refreshed.access}` },
      });
    }
  }

  return res;
}

function mediaUrl(path) {
  if (!path) return '';
  if (path.startsWith('http')) return path;
  return API_BASE_URL + path;
}

function getCroppedImg(imageSrc, pixelCrop) {
  return new Promise((resolve, reject) => {
    const image = new window.Image();
    image.src = imageSrc;
    image.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width = pixelCrop.width;
      canvas.height = pixelCrop.height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(
        image,
        pixelCrop.x, pixelCrop.y,
        pixelCrop.width, pixelCrop.height,
        0, 0,
        pixelCrop.width, pixelCrop.height
      );
      canvas.toBlob((blob) => {
        if (blob) {
          resolve(blob);
        } else {
          const dataUrl = canvas.toDataURL('image/jpeg', 0.9);
          fetch(dataUrl).then(r => r.blob()).then(resolve).catch(reject);
        }
      }, 'image/jpeg', 0.9);
    };
    image.onerror = reject;
  });
}

function ImageCropperModal({ imageSrc, aspect, onCrop, onCancel }) {
  const [crop, setCrop] = useState({ x: 0, y: 0 });
  const [zoom, setZoom] = useState(1);
  const [selectedRatio, setSelectedRatio] = useState(aspect);
  const [autoCropSize, setAutoCropSize] = useState(null);
  const [freeCropSize, setFreeCropSize] = useState(null);
  const [resizeStart, setResizeStart] = useState(null);
  const [croppedAreaPixels, setCroppedAreaPixels] = useState(null);
  const bodyRef = useRef(null);

  const onCropComplete = useCallback((_, area) => {
    setCroppedAreaPixels(area);
  }, []);

  const onCropSizeChange = useCallback((size) => {
    setAutoCropSize((prev) => (
      prev && Math.abs(prev.width - size.width) < 1 && Math.abs(prev.height - size.height) < 1
        ? prev
        : size
    ));
  }, []);

  useEffect(() => {
    if (selectedRatio !== null || freeCropSize || !autoCropSize) return;
    setFreeCropSize({ ...autoCropSize });
  }, [selectedRatio, freeCropSize, autoCropSize]);

  useEffect(() => {
    if (!resizeStart) return undefined;
    const handleMove = (e) => {
      const rect = bodyRef.current?.getBoundingClientRect();
      if (!rect) return;
      const width = Math.min(rect.width, Math.max(60, resizeStart.width + (e.clientX - resizeStart.x) * 2));
      const height = Math.min(rect.height, Math.max(60, resizeStart.height + (e.clientY - resizeStart.y) * 2));
      setFreeCropSize({ width, height });
    };
    const handleUp = () => setResizeStart(null);
    window.addEventListener('pointermove', handleMove);
    window.addEventListener('pointerup', handleUp);
    return () => {
      window.removeEventListener('pointermove', handleMove);
      window.removeEventListener('pointerup', handleUp);
    };
  }, [resizeStart]);

  const handleRatioChange = (value) => {
    setSelectedRatio(value);
    setFreeCropSize(value === null && autoCropSize ? { ...autoCropSize } : null);
    setCrop({ x: 0, y: 0 });
    setZoom(1);
  };

  const handleCrop = async (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (!croppedAreaPixels) return;
    try {
      const blob = await getCroppedImg(imageSrc, croppedAreaPixels);
      onCrop(blob);
    } catch (err) {
      console.error('Crop failed', err);
    }
  };

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="cropper-modal" onClick={(e) => e.stopPropagation()}>
        <div className="cropper-head">
          <h3>Crop Image</h3>
          <button type="button" className="profile-close" onClick={onCancel}><X size={18} /></button>
        </div>
        <div className="cropper-body" ref={bodyRef}>
          <Cropper
            image={imageSrc}
            crop={crop}
            zoom={zoom}
            aspect={selectedRatio ?? aspect}
            cropSize={freeCropSize || undefined}
            onCropSizeChange={onCropSizeChange}
            onCropChange={setCrop}
            onZoomChange={setZoom}
            onCropComplete={onCropComplete}
          />
          {freeCropSize && (
            <div
              className="crop-free-frame"
              style={{ width: freeCropSize.width, height: freeCropSize.height }}
            >
              <button
                type="button"
                className="crop-free-handle"
                aria-label="Resize crop area"
                onPointerDown={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  setResizeStart({ x: e.clientX, y: e.clientY, width: freeCropSize.width, height: freeCropSize.height });
                }}
              />
            </div>
          )}
        </div>
        <div className="cropper-controls">
          <div className="cropper-ratio-row">
            {CROP_PRESETS.map((preset) => (
              <button
                key={preset.label}
                type="button"
                className={`crop-ratio-button ${selectedRatio === preset.value ? 'active' : ''}`}
                onClick={() => handleRatioChange(preset.value)}
              >
                {preset.label}
              </button>
            ))}
          </div>
          <div className="cropper-zoom-row">
            <label>Zoom</label>
            <input
              type="range"
              min={1}
              max={3}
              step={0.1}
              value={zoom}
              onChange={(e) => setZoom(Number(e.target.value))}
            />
          </div>
          <div className="cropper-actions">
            <button type="button" className="small-button" onClick={onCancel}>Cancel</button>
            <button type="button" className="primary-button" onClick={handleCrop}>
              <Crop size={16} /> Apply
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function exportToExcel(rows, columns, filename) {
  const data = rows.map((row) => {
    const obj = {};
    columns.forEach((col) => { obj[col.header] = col.getValue(row); });
    return obj;
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = columns.map(() => ({ wch: 22 }));
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, `${filename}.xlsx`);
  window.__showToast?.(`${filename}.xlsx downloaded successfully!`);
}

function Toast({ message, onClose }) {
  useEffect(() => {
    const t = setTimeout(onClose, 3000);
    return () => clearTimeout(t);
  }, [onClose]);

  return (
    <div className="toast-popup" onClick={onClose}>
      <CheckCircle2 size={18} />
      <span>{message}</span>
    </div>
  );
}

function EmptyState({ icon: Icon, title, description, actionLabel, onAction }) {
  return (
    <div className="empty-state">
      <div className="empty-state-icon">
        <Icon size={48} strokeWidth={1.2} />
      </div>
      <h3 className="empty-state-title">{title}</h3>
      <p className="empty-state-desc">{description}</p>
      {actionLabel && onAction && (
        <button className="primary-button empty-state-cta" onClick={onAction}>
          <Plus size={16} /> {actionLabel}
        </button>
      )}
    </div>
  );
}

const defaultNavItems = [
  { icon: Home, label: 'Dashboard' },
  { icon: Store, label: 'Branches' },
  { icon: Grid2X2Plus, label: 'Offers' },
  { icon: Image, label: 'Banners' },
  { icon: Bell, label: 'Notifications' },
  { icon: Users, label: 'Customers' },
  { icon: ClipboardList, label: 'Forms' },
  { icon: Users, label: 'Users' },
  { icon: Clock3, label: 'Activity' },
];

const iconMap = { Home, Store, Grid2X2Plus, Image, Bell, ClipboardList, Users };

function loadNavItems() {
  try {
    const saved = localStorage.getItem('sidebar-order');
    if (!saved) return defaultNavItems;
    const order = JSON.parse(saved);
    if (!Array.isArray(order) || order.length !== defaultNavItems.length) return defaultNavItems;
    return order.map((label) => {
      const found = defaultNavItems.find((item) => item.label === label);
      return found || defaultNavItems.find((d) => d.label === label);
    }).filter(Boolean);
  } catch {
    return defaultNavItems;
  }
}

function saveNavItems(items) {
  localStorage.setItem('sidebar-order', JSON.stringify(items.map((i) => i.label)));
}

const quickActions = [LogOut, Search, Target];

function LoginPage({ onLogin }) {
  const [phone, setPhone] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const handleSubmit = async (event) => {
    event.preventDefault();
    setLoading(true);
    setError('');

    try {
      const res = await fetch(`${API_BASE_URL}/api/auth/admin-login/`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ phone, password }),
      });

      const data = await res.json();

      if (!res.ok) {
        throw new Error(data.detail || data.message || 'Login failed');
      }

      onLogin(data);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <main className="login-page">
      <section className="login-shell">
        <form className="login-card" onSubmit={handleSubmit}>
          <div className="login-card-head">
            <img src="/assets/logo-hyp.png" alt="Hyper City" className="login-logo" />
            <div>
              <p>Welcome back</p>
              <h2>Sign in to dashboard</h2>
            </div>
          </div>

          {error && <div className="login-error">{error}</div>}

          <label className="field">
            <span>Phone number</span>
            <div>
              <Mail size={18} />
              <input type="tel" value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="Enter phone number" aria-label="Phone number" />
            </div>
          </label>

          <label className="field">
            <span>Password</span>
            <div>
              <LockKeyhole size={18} />
              <input type={showPassword ? 'text' : 'password'} value={password} onChange={(e) => setPassword(e.target.value)} placeholder="Enter password" aria-label="Password" />
              <button type="button" className="field-eye" onClick={() => setShowPassword((p) => !p)} aria-label={showPassword ? 'Hide password' : 'Show password'}>
                {showPassword ? <EyeOff size={18} /> : <Eye size={18} />}
              </button>
            </div>
          </label>

          <div className="login-options">
            <label>
              <input type="checkbox" defaultChecked />
              <span>Remember me</span>
            </label>
            <button type="button">Forgot password?</button>
          </div>

          <button className="login-submit" type="submit" disabled={loading}>
            <LogIn size={18} />
            {loading ? 'Signing in...' : 'Login'}
          </button>
        </form>

        <div className="login-visual">
          <img src="/assets/login_graphics.png" alt="" className="login-graphics" />
        </div>
      </section>
    </main>
  );
}

function Sidebar({ activePage, onNavigate, navItems, editMode, onToggleEdit, onReorder }) {
  const dragIndex = useRef(null);

  function handleDragStart(e, index) {
    dragIndex.current = index;
    e.dataTransfer.effectAllowed = 'move';
    e.currentTarget.classList.add('dragging');
  }

  function handleDragEnd(e) {
    e.currentTarget.classList.remove('dragging');
    document.querySelectorAll('.nav-item.drag-over').forEach((el) => el.classList.remove('drag-over'));
  }

  function handleDragOver(e, index) {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    const items = document.querySelectorAll('.nav-list .nav-item');
    items.forEach((el) => el.classList.remove('drag-over'));
    if (dragIndex.current !== null && dragIndex.current !== index) {
      items[index].classList.add('drag-over');
    }
  }

  function handleDrop(e, index) {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-over');
    if (dragIndex.current === null || dragIndex.current === index) return;
    const newItems = [...navItems];
    const [moved] = newItems.splice(dragIndex.current, 1);
    newItems.splice(index, 0, moved);
    dragIndex.current = null;
    onReorder(newItems);
  }

  return (
    <aside className={`sidebar ${editMode ? 'editing' : ''}`}>
      <div className="brand">
        <img src="/assets/logo-hyp.png" alt="Hyper City" className="brand-logo" />
        <span>Hyper City</span>
        <ChevronDown size={13} />
      </div>
      <nav className="nav-list" aria-label="Main navigation">
        {navItems.map(({ icon: Icon, label }, index) => (
          <button
            className={`nav-item ${activePage === label && !editMode ? 'active' : ''}`}
            key={label}
            onClick={() => { if (!editMode) onNavigate(label); }}
            draggable={editMode}
            onDragStart={(e) => handleDragStart(e, index)}
            onDragEnd={handleDragEnd}
            onDragOver={(e) => handleDragOver(e, index)}
            onDrop={(e) => handleDrop(e, index)}
          >
            {editMode && <GripVertical size={14} className="drag-handle" />}
            <Icon size={18} />
            <span>{label}</span>
          </button>
        ))}
      </nav>
      {editMode && (
        <button className="edit-done-btn" onClick={onToggleEdit}>
          <CheckCircle2 size={15} />
          Done
        </button>
      )}
    </aside>
  );
}

function ProfilePanel({ authData, onClose }) {
  const [profile, setProfile] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    apiFetch(`/api/profile/`)
      .then((res) => res.json())
      .then((data) => {
        setProfile(data);
        setLoading(false);
      })
      .catch(() => setLoading(false));
  }, []);

  const user = profile || authData.user;

  return (
    <div className="profile-overlay" onClick={onClose}>
      <aside className="profile-panel" onClick={(e) => e.stopPropagation()}>
        <header className="profile-head">
          <strong>My Profile</strong>
          <button className="profile-close" onClick={onClose}><X size={18} /></button>
        </header>

        {loading && !profile ? (
          <p className="profile-loading">Loading...</p>
        ) : (
          <>
            <div className="profile-avatar-row">
              <span className="profile-avatar">{user.name?.[0] || 'U'}</span>
              <div>
                <strong>{user.name}</strong>
                <span>{user.role?.replace('_', ' ')}</span>
              </div>
            </div>

            <div className="profile-details">
              <div className="profile-field">
                <Phone size={16} />
                <span>{user.phone}</span>
              </div>
              <div className="profile-field">
                <Mail size={16} />
                <span>{user.email || '—'}</span>
              </div>
              <div className="profile-field">
                <Users size={16} />
                <span>ID: {user.id}</span>
              </div>
              {user.managed_branch && (
                <div className="profile-field">
                  <Store size={16} />
                  <span>Branch: {user.managed_branch}</span>
                </div>
              )}
            </div>

            {profile && profile.additional_info && (
              <div className="profile-meta">
                {Object.entries(profile.additional_info).map(([key, val]) => (
                  <div className="profile-meta-row" key={key}>
                    <em>{key}</em>
                    <span>{String(val)}</span>
                  </div>
                ))}
              </div>
            )}
          </>
        )}
      </aside>
    </div>
  );
}

function GlobalSearch({ onClose, onNavigate }) {
  const [query, setQuery] = useState('');
  const [results, setResults] = useState({ branches: [], offers: [], banners: [], staff: [] });
  const [loading, setLoading] = useState(false);
  const inputRef = useRef(null);

  useEffect(() => {
    inputRef.current?.focus();
  }, []);

  useEffect(() => {
    if (!query.trim()) {
      setResults({ branches: [], offers: [], banners: [], staff: [] });
      return;
    }

    const timer = setTimeout(() => {
      setLoading(true);
      const params = new URLSearchParams({ search: query, page_size: 10 });

      Promise.allSettled([
        apiFetch(`/api/branches/?${params}`).then(r => r.json()),
        apiFetch(`/api/offers/?${params}`).then(r => r.json()),
        apiFetch(`/api/banners/?${params}`).then(r => r.json()),
        apiFetch(`/api/staff/?${params}`).then(r => r.json()),
      ]).then(([branchesRes, offersRes, bannersRes, staffRes]) => {
        setResults({
          branches: branchesRes.status === 'fulfilled' ? branchesRes.value.results || [] : [],
          offers: offersRes.status === 'fulfilled' ? offersRes.value.results || [] : [],
          banners: bannersRes.status === 'fulfilled' ? bannersRes.value.results || [] : [],
          staff: staffRes.status === 'fulfilled' ? staffRes.value.results || [] : [],
        });
        setLoading(false);
      });
    }, 300);

    return () => clearTimeout(timer);
  }, [query]);

  const hasResults = results.branches.length || results.offers.length || results.banners.length || results.staff.length;

  const handleResultClick = (page, options) => {
    onNavigate(page, options);
    onClose();
  };

  return (
    <div className="search-overlay" onClick={onClose}>
      <div className="search-popup" onClick={(e) => e.stopPropagation()}>
        <div className="search-popup-input">
          <Search size={18} />
          <input
            ref={inputRef}
            type="text"
            placeholder="Search branches, offers, banners, staff..."
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
          <button className="search-close-btn" onClick={onClose}><X size={16} /></button>
        </div>

        <div className="search-popup-results">
          {loading && <p className="search-status">Searching...</p>}

          {!loading && query && !hasResults && (
            <p className="search-status">No results found for "{query}"</p>
          )}

          {!loading && !query && (
            <p className="search-status">Type to search across all sections...</p>
          )}

          {results.branches.length > 0 && (
            <div className="search-group">
              <button className="search-group-heading" onClick={() => handleResultClick('Branches')}>
                <Store size={14} /> Branches ({results.branches.length})
              </button>
              {results.branches.map((b) => (
                <button className="search-result-item" key={`b-${b.id}`} onClick={() => handleResultClick('Branches', { branch: { id: b.id, name: b.name } })}>
                  <span className="search-result-avatar">{b.name?.[0] || 'B'}</span>
                  <div>
                    <strong>{b.name}</strong>
                    <span>{b.address}</span>
                  </div>
                </button>
              ))}
            </div>
          )}

          {results.offers.length > 0 && (
            <div className="search-group">
              <button className="search-group-heading" onClick={() => handleResultClick('Offers')}>
                <Grid2X2Plus size={14} /> Offers ({results.offers.length})
              </button>
              {results.offers.map((o) => (
                <button className="search-result-item" key={`o-${o.id}`} onClick={() => handleResultClick('Offers')}>
                  {o.image ? <img src={mediaUrl(o.image)} alt="" className="search-result-thumb" /> : <span className="search-result-avatar">{o.title?.[0] || 'O'}</span>}
                  <div>
                    <strong>{o.title}</strong>
                    <span>{o.discount_type === 'PERCENT' ? `${o.discount_value}% OFF` : `₹${o.discount_value} OFF`}</span>
                  </div>
                </button>
              ))}
            </div>
          )}

          {results.banners.length > 0 && (
            <div className="search-group">
              <button className="search-group-heading" onClick={() => handleResultClick('Banners')}>
                <Image size={14} /> Banners ({results.banners.length})
              </button>
              {results.banners.map((b) => (
                <button className="search-result-item" key={`bn-${b.id}`} onClick={() => handleResultClick('Banners')}>
                  {b.image ? <img src={mediaUrl(b.image)} alt="" className="search-result-thumb" /> : <span className="search-result-avatar">{b.title?.[0] || 'B'}</span>}
                  <div>
                    <strong>{b.title}</strong>
                    <span>{b.link_url || 'No link'}</span>
                  </div>
                </button>
              ))}
            </div>
          )}

          {results.staff.length > 0 && (
            <div className="search-group">
              <button className="search-group-heading" onClick={() => handleResultClick('Users')}>
                <Users size={14} /> Staff ({results.staff.length})
              </button>
              {results.staff.map((s) => (
                <button className="search-result-item" key={`s-${s.id}`} onClick={() => handleResultClick('Users')}>
                  <span className="search-result-avatar">{s.name?.[0] || 'S'}</span>
                  <div>
                    <strong>{s.name}</strong>
                    <span>{s.role === 'BRANCH_MGR' ? 'Branch Manager' : 'Marketing'}</span>
                  </div>
                </button>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function Rail({ authData, onOpenSearch, onLogout, editMode, onToggleEdit }) {
  const [showProfile, setShowProfile] = useState(false);
  const [confirmLogout, setConfirmLogout] = useState(false);

  return (
    <>
      <aside className="rail">
        <div className="rail-actions">
          {quickActions.map((Icon, index) => (
            <button
              className={`rail-button ${index === 0 ? 'primary' : ''} ${index === 2 && editMode ? 'active-edit' : ''}`}
              key={index}
              aria-label={`Quick action ${index + 1}`}
              onClick={() => { if (index === 0) setConfirmLogout(true); else if (index === 1) onOpenSearch(); else if (index === 2) onToggleEdit(); }}
            >
              <Icon size={index === 0 ? 22 : 18} />
            </button>
          ))}
        </div>
        <button className="rail-logo" onClick={() => setShowProfile(true)} aria-label="Profile">
          <span>{authData.user.name?.[0] || 'U'}</span>
        </button>
      </aside>

      {confirmLogout && (
        <div className="modal-overlay" onClick={() => setConfirmLogout(false)}>
          <div className="modal-card confirm-card logout-confirm" onClick={(e) => e.stopPropagation()}>
            <div className="logout-confirm-icon">
              <LogOut size={24} />
            </div>
            <h3>Confirm Logout</h3>
            <p>Are you sure you want to log out of the dashboard?</p>
            <div className="modal-actions">
              <button className="small-button" onClick={() => setConfirmLogout(false)}>Cancel</button>
              <button className="danger-button" onClick={onLogout}>Logout</button>
            </div>
          </div>
        </div>
      )}

      {showProfile && <ProfilePanel authData={authData} onClose={() => setShowProfile(false)} />}
    </>
  );
}

function MetricCard({ icon: Icon, label, value, delta, tone }) {
  return (
    <article className="metric-card">
      <div className={`metric-icon ${tone}`}>
        <Icon size={23} />
      </div>
      <div className="metric-body">
        <p>{label}</p>
        <span>{delta}</span>
      </div>
      <strong className="metric-value">{value}</strong>
    </article>
  );
}

function LineOverview({ overview, periodLabel }) {
  const labels = overview?.labels || [];
  const active = overview?.active || [];
  const posts = overview?.posts || [];

  const pointCount = Math.max(labels.length, active.length, posts.length, 1);
  const xs = Array.from({ length: pointCount }, (_, i) => (
    pointCount === 1 ? 332 : 58 + (510 * i) / (pointCount - 1)
  ));
  const baseY = 260;
  const max = Math.max(1, ...active, ...posts);

  const yFor = (v) => (baseY - (v / max) * 200).toFixed(2);

  const toPath = (series) => series
    .map((v, i) => `${i === 0 ? 'M' : 'L'} ${xs[i]} ${yFor(v)}`)
    .join(' ');

  const activePath = toPath(active);
  const postsPath = toPath(posts);
  const gridValues = [0.25, 0.5, 0.75, 1].map((f) => max * f);

  return (
    <section className="panel chart-panel">
      <div className="panel-header">
        <h2>Offers Overview</h2>
        <span className="pill-button">{periodLabel}</span>
      </div>
      <div className="legend">
        <span><i className="active-line" />Active Offers</span>
        <span><i className="posts-line" />Offer Posts</span>
      </div>
      <div className="chart-wrap">
        <svg viewBox="0 0 650 300" role="img" aria-label="Offers line chart">
          {gridValues.map((gv, index) => (
            <g key={index}>
              <text x="20" y={baseY - (gv / max) * 200 + 5}>{Math.round(gv)}</text>
              <line x1="58" x2="606" y1={baseY - (gv / max) * 200} y2={baseY - (gv / max) * 200} />
            </g>
          ))}
          <text x="25" y="265">0</text>
          <line className="axis" x1="58" x2="606" y1="260" y2="260" />
          <path d={activePath} className="blue-path" />
          <path d={postsPath} className="purple-path" />
          {xs.map((x, i) => (
            <g key={x}>
              <line className="tick" x1={x} x2={x} y1="260" y2="270" />
              <text className="date" x={x - 23} y="290">{labels[i] || `Day ${i + 1}`}</text>
            </g>
          ))}
          {[
            ...active.map((v, i) => ({ v, i, color: 'blue' })),
            ...posts.map((v, i) => ({ v, i, color: 'purple' })),
          ].map(({ v, i, color }, idx) => (
            <circle key={idx} className={color} cx={xs[i]} cy={yFor(v)} r="5" />
          ))}
        </svg>
      </div>
    </section>
  );
}

function BranchPanel({ topBranches = [], onViewAll }) {
  const tones = ['indigo', 'blue', 'orange', 'purple', 'teal'];
  return (
    <section className="panel branch-panel">
      <div className="panel-header">
        <h2>Top Performing Branches</h2>
        <button type="button" className="small-button" onClick={onViewAll}>View All</button>
      </div>
      <div className="branch-list">
        {topBranches.map((b, index) => (
          <div className="branch-row" key={b.name}>
            <b>{index + 1}</b>
            <span className={`avatar ${tones[index % tones.length]}`}>{b.initials}</span>
            <strong>{b.name}</strong>
            <em>{b.offer_count} Offers</em>
            <span className="growth">{b.growth >= 0 ? '+' : ''}{b.growth}%</span>
          </div>
        ))}
      </div>
    </section>
  );
}

function Thumbnail({ image, title }) {
  if (image) {
    return <img src={mediaUrl(image)} alt="" className="post-thumb-img" />;
  }
  return (
    <div className="thumb thumb-generic">
      <span>{title?.[0] || 'O'}</span>
    </div>
  );
}

function PostsPanel({ recentPosts = [], onViewAll }) {
  return (
    <section className="panel posts-panel">
      <div className="panel-header">
        <h2>Recent Offer Posts</h2>
        <button type="button" className="small-button" onClick={onViewAll}>View All</button>
      </div>
      <div className="post-list">
        {recentPosts.map((post) => {
          const statusClass = post.status === 'Active' ? '' : post.status === 'Scheduled' ? 'scheduled' : 'pending';
          return (
            <article className="post-row" key={post.title}>
              <Thumbnail image={post.image} title={post.title} />
              <div className="post-copy">
                <strong>{post.title}</strong>
                <span>{post.place}</span>
              </div>
              <span className={`status ${statusClass}`}>{post.status}</span>
              <time>{post.time_label}</time>
              <button aria-label={`More actions for ${post.title}`}>
                <MoreVertical size={17} />
              </button>
            </article>
          );
        })}
      </div>
    </section>
  );
}

function DonutPanel({ distribution = [] }) {
  let cumulative = 0;
  const stops = distribution.map((d) => {
    const from = cumulative;
    cumulative += d.pct;
    return `${d.color} ${from}% ${cumulative}%`;
  });
  const total = distribution.reduce((sum, d) => sum + d.value, 0);
  const background = cumulative > 0 ? `conic-gradient(${stops.join(', ')})` : '#eef1f8';

  return (
    <section className="panel donut-panel">
      <h2>Offer Status Distribution</h2>
      <div className="donut-content">
        <div className="donut" aria-label="Offer status donut chart" style={{ background }}>
          <div className="donut-hole">
            <strong>{total}</strong>
            <span>Total Offers</span>
          </div>
        </div>
        <div className="status-list">
          {distribution.map((d) => (
            <div className="status-row" key={d.label}>
              <i style={{ backgroundColor: d.color }} />
              <strong>{d.label}</strong>
              <span>{d.value}</span>
              <em>({d.pct}%)</em>
            </div>
          ))}
        </div>
      </div>
    </section>
  );
}

function BulkActionBar({ count, onDelete, onSetActive, onSetInactive, onClear, busy }) {
  return (
    <div className="bulk-action-bar">
      <div className="bulk-left">
        <button className="bulk-clear" onClick={onClear}>
          <X size={16} />
        </button>
        <span className="bulk-count"><strong>{count}</strong> selected</span>
      </div>
      <div className="bulk-right">
        <button className="bulk-btn bulk-active" onClick={onSetActive} disabled={busy}>
          <Eye size={15} /> Set Active
        </button>
        <button className="bulk-btn bulk-inactive" onClick={onSetInactive} disabled={busy}>
          <EyeOff size={15} /> Set Inactive
        </button>
        <button className="bulk-btn bulk-delete" onClick={onDelete} disabled={busy}>
          <Trash2 size={15} /> Delete
        </button>
      </div>
    </div>
  );
}

function BranchesPage({ selectedBranch, onClearSelectedBranch }) {
  const [branches, setBranches] = useState([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(1);
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const [loading, setLoading] = useState(true);
  const [modal, setModal] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);
  const [selectedIds, setSelectedIds] = useState([]);
  const [bulkBusy, setBulkBusy] = useState(false);
  const pageSize = 10;

  const fetchBranches = useCallback(async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, page_size: pageSize });
      if (search) params.set('search', search);
      if (selectedBranch?.id) params.set('branch_id', selectedBranch.id);
      if (statusFilter === 'active') params.set('is_active', 'true');
      if (statusFilter === 'inactive') params.set('is_active', 'false');
      const res = await apiFetch(`/api/branches/?${params}`);
      const data = await res.json();
      setBranches(data.results || []);
      setTotal(data.count || 0);
    } catch (err) {
      console.error('Failed to fetch branches', err);
    } finally {
      setLoading(false);
    }
  }, [page, search, statusFilter, selectedBranch?.id]);

  useEffect(() => { fetchBranches(); }, [fetchBranches]);

  useEffect(() => {
    setPage(1);
  }, [selectedBranch?.id]);

  const filteredBranches = branches;

  const handleDelete = async () => {
    if (!deleteTarget) return;
    try {
      await apiFetch(`/api/branches/${deleteTarget.id}/`, {
        method: 'DELETE',
      });
      setDeleteTarget(null);
      fetchBranches();
    } catch (err) {
      console.error('Failed to delete branch', err);
    }
  };

  const toggleSelect = (id) => {
    setSelectedIds((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  };

  const selectAll = () => {
    const allIds = filteredBranches.map((b) => b.id);
    setSelectedIds((prev) => prev.length === allIds.length ? [] : allIds);
  };

  const handleBulkDelete = async () => {
    if (!selectedIds.length) return;
    setBulkBusy(true);
    await Promise.allSettled(
      selectedIds.map((id) =>
        apiFetch(`/api/branches/${id}/`, {
          method: 'DELETE',
        })
      )
    );
    setSelectedIds([]);
    setBulkBusy(false);
    fetchBranches();
  };

  const handleBulkSetActive = async () => {
    if (!selectedIds.length) return;
    setBulkBusy(true);
    await Promise.allSettled(
      selectedIds.map((id) =>
        apiFetch(`/api/branches/${id}/`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ is_active: true }),
        })
      )
    );
    setSelectedIds([]);
    setBulkBusy(false);
    setStatusFilter('all');
    setPage(1);
    fetchBranches();
  };

  const handleBulkSetInactive = async () => {
    if (!selectedIds.length) return;
    setBulkBusy(true);
    await Promise.allSettled(
      selectedIds.map((id) =>
        apiFetch(`/api/branches/${id}/`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ is_active: false }),
        })
      )
    );
    setSelectedIds([]);
    setBulkBusy(false);
    setStatusFilter('all');
    setPage(1);
    fetchBranches();
  };

  const totalPages = Math.ceil(total / pageSize);

  const filterTabs = [
    { key: 'all', label: 'All' },
    { key: 'active', label: 'Active' },
    { key: 'inactive', label: 'Inactive' },
  ];

  return (
    <>
      <div className="hero-row">
        <div>
          <p>Manage Locations</p>
          <h2>Branches</h2>
        </div>
        <div className="filters">
          {selectedBranch && (
            <button type="button" className="small-button" onClick={onClearSelectedBranch} title="Show all branches">
              {selectedBranch.name} <X size={15} />
            </button>
          )}
          <div className="search-field">
            <Search size={17} />
            <input type="text" placeholder="Search branches..." value={search} onChange={(e) => { setSearch(e.target.value); setPage(1); }} />
          </div>
          <button className="primary-button" onClick={() => setModal({ mode: 'create' })}>
            <Plus size={18} /> Add Branch
          </button>
          <button className="export-button" onClick={() => exportToExcel(filteredBranches, [
            { header: 'Name', getValue: (r) => r.name },
            { header: 'Address', getValue: (r) => r.address },
            { header: 'Phone', getValue: (r) => r.phone },
            { header: 'Status', getValue: (r) => r.is_active ? 'Active' : 'Inactive' },
          ], 'branches')}>
            <Download size={16} /> Export
          </button>
        </div>
      </div>

      <section className="panel branches-panel">
        <div className="filter-tabs">
          {filterTabs.map((t) => (
            <button key={t.key} className={`filter-tab ${statusFilter === t.key ? 'active' : ''}`} onClick={() => { setStatusFilter(t.key); setPage(1); setSelectedIds([]); }}>
              {t.label}
            </button>
          ))}
        </div>

        {loading ? (
          <p className="table-loading">Loading branches...</p>
        ) : filteredBranches.length === 0 ? (
          search || statusFilter !== 'all' ? (
            <EmptyState icon={Search} title="No matching branches" description={`No branches match your current filters. Try adjusting your search or filters.`} />
          ) : (
            <EmptyState icon={Store} title="No branches yet" description="Get started by adding your first branch location." actionLabel="Add Branch" onAction={() => setModal({ mode: 'create' })} />
          )
        ) : (
          <div className="branches-table">
            <div className="branches-table-head">
              <label className="bulk-checkbox-cell">
                <input type="checkbox" checked={selectedIds.length === filteredBranches.length && filteredBranches.length > 0} onChange={selectAll} />
              </label>
              <span />
              <span>Name</span>
              <span>Address</span>
              <span>Phone</span>
              <span>Status</span>
              <span>Actions</span>
            </div>
            {filteredBranches.map((b) => (
              <div className="branches-table-row" key={b.id}>
                <label className="bulk-checkbox-cell">
                  <input type="checkbox" checked={selectedIds.includes(b.id)} onChange={() => toggleSelect(b.id)} />
                </label>
                <span className="branch-avatar">{b.name?.[0] || 'B'}</span>
                <strong>{b.name}</strong>
                <span className="branch-address">{b.address}</span>
                <span className="branch-phone">{b.phone}</span>
                <span className="branch-status-cell">
                  <span className={`status ${!b.is_active ? 'pending' : ''}`}>{b.is_active ? 'Active' : 'Inactive'}</span>
                </span>
                <span className="branch-actions">
                  <button onClick={() => setModal({ mode: 'edit', branch: b })} aria-label="Edit"><Edit3 size={16} /></button>
                  <button onClick={() => setDeleteTarget(b)} aria-label="Delete"><Trash2 size={16} /></button>
                </span>
              </div>
            ))}
          </div>
        )}

        {totalPages > 1 && (
          <div className="pagination">
            <button disabled={page <= 1} onClick={() => setPage((p) => p - 1)}><ChevronLeft size={16} /> Prev</button>
            <span>Page {page} of {totalPages}</span>
            <button disabled={page >= totalPages} onClick={() => setPage((p) => p + 1)}>Next <ChevronRight size={16} /></button>
          </div>
        )}
      </section>

      {modal && (
        <BranchFormModal
          mode={modal.mode}
          branch={modal.branch}
          onClose={() => setModal(null)}
          onSaved={() => { setModal(null); fetchBranches(); }}
        />
      )}

      {deleteTarget && (
        <div className="modal-overlay" onClick={() => setDeleteTarget(null)}>
          <div className="modal-card confirm-card" onClick={(e) => e.stopPropagation()}>
            <h3>Delete Branch</h3>
            <p>Are you sure you want to delete <strong>{deleteTarget.name}</strong>?</p>
            <div className="modal-actions">
              <button className="small-button" onClick={() => setDeleteTarget(null)}>Cancel</button>
              <button className="danger-button" onClick={handleDelete}>Delete</button>
            </div>
          </div>
        </div>
      )}

      {selectedIds.length > 0 && (
        <BulkActionBar
          count={selectedIds.length}
          onDelete={handleBulkDelete}
          onSetActive={handleBulkSetActive}
          onSetInactive={handleBulkSetInactive}
          onClear={() => setSelectedIds([])}
          busy={bulkBusy}
        />
      )}
    </>
  );
}

function BranchFormModal({ mode, branch, onClose, onSaved }) {
  const [form, setForm] = useState({
    name: branch?.name || '',
    address: branch?.address || '',
    phone: branch?.phone || '',
    is_active: branch?.is_active ?? true,
  });
  const [saving, setSaving] = useState(false);

  const handleChange = (field) => (e) => {
    const value = field === 'is_active' ? e.target.checked : e.target.value;
    setForm((prev) => ({ ...prev, [field]: value }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setSaving(true);
    try {
      const path = mode === 'create'
        ? '/api/branches/'
        : `/api/branches/${branch.id}/`;
      const method = mode === 'create' ? 'POST' : 'PUT';

      const res = await apiFetch(path, {
        method,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(form),
      });

      if (!res.ok) throw new Error('Failed to save branch');
      onSaved();
    } catch (err) {
      console.error(err);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card" onClick={(e) => e.stopPropagation()} onSubmit={handleSubmit}>
        <div className="modal-head">
          <h3>{mode === 'create' ? 'Add Branch' : 'Edit Branch'}</h3>
          <button type="button" className="profile-close" onClick={onClose}><X size={18} /></button>
        </div>

        <label className="modal-field">
          <span>Name</span>
          <input type="text" value={form.name} onChange={handleChange('name')} required />
        </label>

        <label className="modal-field">
          <span>Address</span>
          <input type="text" value={form.address} onChange={handleChange('address')} required />
        </label>

        <label className="modal-field">
          <span>Phone</span>
          <input type="tel" inputMode="numeric" pattern="[0-9]{10}" minLength="10" maxLength="10" title="Enter a valid 10-digit phone number" value={form.phone} onChange={handleChange('phone')} required />
        </label>

        <label className="modal-check">
          <input type="checkbox" checked={form.is_active} onChange={handleChange('is_active')} />
          <span>Active</span>
        </label>

        <div className="modal-actions">
          <button type="button" className="small-button" onClick={onClose}>Cancel</button>
          <button type="submit" className="primary-button modal-submit" disabled={saving}>
            {saving ? 'Saving...' : mode === 'create' ? 'Create' : 'Update'}
          </button>
        </div>
      </form>
    </div>
  );
}

const dashboardIcons = {
  total_branches: Store,
  active_offers: PackageOpen,
  offer_posts: CalendarDays,
};

const hiddenMetrics = new Set(['total_views']);

function dateInputValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function thisMonthRange() {
  const start = new Date();
  start.setDate(1);
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setMonth(end.getMonth() + 1);
  end.setDate(0);
  return { start: dateInputValue(start), end: dateInputValue(end) };
}

function dateRangeLabel(range) {
  const start = new Date(`${range.start}T00:00:00`);
  const end = new Date(`${range.end}T00:00:00`);
  const opts = { day: 'numeric', month: 'short', year: 'numeric' };
  return `${start.toLocaleDateString('en-IN', opts)} - ${end.toLocaleDateString('en-IN', opts)}`;
}

function DashboardPage({ authData, onNavigate }) {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [isDatePickerOpen, setIsDatePickerOpen] = useState(false);
  const [appliedRange, setAppliedRange] = useState(() => thisMonthRange());
  const [draftRange, setDraftRange] = useState(() => thisMonthRange());
  const [dateError, setDateError] = useState('');

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError('');
    const params = new URLSearchParams({ start_date: appliedRange.start, end_date: appliedRange.end });
    apiFetch(`/api/dashboard/?${params}`)
      .then(async (res) => {
        const json = await res.json();
        if (!res.ok) throw new Error(json.detail || 'Failed to load dashboard data');
        return json;
      })
      .then((json) => {
        if (cancelled) return;
        setData(json);
        setLoading(false);
      })
      .catch((requestError) => {
        if (cancelled) return;
        setError(requestError.message || 'Failed to load dashboard data');
        setLoading(false);
      });
    return () => { cancelled = true; };
  }, [appliedRange]);

  const applyDateRange = () => {
    if (!draftRange.start || !draftRange.end) {
      setDateError('Choose both a start and end date.');
      return;
    }
    if (draftRange.start > draftRange.end) {
      setDateError('The end date must be on or after the start date.');
      return;
    }
    setDateError('');
    setAppliedRange(draftRange);
    setIsDatePickerOpen(false);
  };

  const resetToThisMonth = () => {
    const range = thisMonthRange();
    setDraftRange(range);
    setAppliedRange(range);
    setDateError('');
    setIsDatePickerOpen(false);
  };

  if (loading) {
    return <p className="table-loading">Loading dashboard...</p>;
  }

  if (error || !data) {
    return <p className="table-loading">{error || 'No dashboard data available'}</p>;
  }

  const metrics = data.metrics
    .filter((m) => !hiddenMetrics.has(m.key))
    .map((m) => ({
      ...m,
      icon: dashboardIcons[m.key] || PackageOpen,
    }));

  return (
    <>
      <div className="hero-row">
        <div>
          <p>Welcome back, {authData.user.name}!👋</p>
          <h2>Manage offers across all branches<br />and grow your business.</h2>
        </div>
        <div className="filters">
          <div className="dashboard-date-filter">
            <button type="button" onClick={() => { setDraftRange(appliedRange); setDateError(''); setIsDatePickerOpen((open) => !open); }} aria-expanded={isDatePickerOpen}>
              <CalendarDays size={18} />{dateRangeLabel(appliedRange)} <ChevronDown size={15} />
            </button>
            {isDatePickerOpen && (
              <div className="date-filter-popover">
                <label>From<input type="date" value={draftRange.start} onChange={(e) => setDraftRange((range) => ({ ...range, start: e.target.value }))} /></label>
                <label>To<input type="date" value={draftRange.end} onChange={(e) => setDraftRange((range) => ({ ...range, end: e.target.value }))} /></label>
                {dateError && <p className="date-filter-error">{dateError}</p>}
                <div className="date-filter-actions">
                  <button type="button" className="small-button" onClick={resetToThisMonth}>This month</button>
                  <button type="button" className="primary-button" onClick={applyDateRange}>Apply</button>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      <section className="metrics" aria-label="Dashboard metrics">
        {metrics.map((metric) => <MetricCard key={metric.key} {...metric} />)}
      </section>

      <section className="dashboard-grid">
        <LineOverview overview={data.overview} periodLabel={dateRangeLabel(appliedRange)} />
        <BranchPanel topBranches={data.top_branches} onViewAll={() => onNavigate('Branches')} />
        <PostsPanel recentPosts={data.recent_posts} onViewAll={() => onNavigate('Offers')} />
        <DonutPanel distribution={data.status_distribution} />
      </section>
    </>
  );
}

function BranchCheckDropdown({ selected, onChange, branches: suppliedBranches }) {
  const [open, setOpen] = useState(false);
  const [fetchedBranches, setFetchedBranches] = useState([]);
  const ref = useRef(null);

  useEffect(() => {
    if (suppliedBranches) return;
    apiFetch(`/api/branches/`)
      .then((r) => r.json())
      .then((data) => setFetchedBranches(data.results || []))
      .catch(() => {});
  }, [suppliedBranches]);

  const branches = suppliedBranches || fetchedBranches;

  useEffect(() => {
    const handler = (e) => {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const toggle = (id) => {
    const next = selected.includes(id) ? selected.filter((x) => x !== id) : [...selected, id];
    onChange(next);
  };

  return (
    <div className="branch-dropdown" ref={ref}>
      <button type="button" className="branch-dropdown-trigger" onClick={() => setOpen((o) => !o)}>
        <span>{selected.length ? `${selected.length} branch${selected.length > 1 ? 'es' : ''} selected` : 'Choose branches'}</span>
        <ChevronDown size={16} />
      </button>
      {open && (
        <div className="branch-dropdown-menu">
          {branches.map((b) => (
            <label key={b.id} className="branch-dropdown-item">
              <input type="checkbox" checked={selected.includes(b.id)} onChange={() => toggle(b.id)} />
              <span>{b.name}</span>
            </label>
          ))}
          {branches.length === 0 && <span className="branch-dropdown-empty">No branches available</span>}
        </div>
      )}
    </div>
  );
}

function OfferFormModal({ mode, offer, onClose, onSaved }) {
  const [form, setForm] = useState({
    title: offer?.title || '',
    description: offer?.description || '',
    image: null,
    start_date: offer?.start_date ? offer.start_date.slice(0, 16) : '',
    end_date: offer?.end_date ? offer.end_date.slice(0, 16) : '',
    terms_conditions: offer?.terms_conditions || '',
    is_active: offer?.is_active ?? true,
  });
  const [selectedBranchIds, setSelectedBranchIds] = useState(offer?.branch_ids || []);
  const [imagePreview, setImagePreview] = useState(offer?.image ? mediaUrl(offer.image) : null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');
  const [cropperSrc, setCropperSrc] = useState(null);

  const handleChange = (field) => (e) => {
    if (field === 'image') {
      const file = e.target.files[0];
      if (file) {
        const preview = URL.createObjectURL(file);
        setCropperSrc(preview);
      }
    } else if (field === 'is_active') {
      setForm((prev) => ({ ...prev, is_active: e.target.checked }));
    } else {
      setForm((prev) => ({ ...prev, [field]: e.target.value }));
    }
  };

  const handleCropped = (blob) => {
    const file = new File([blob], 'cropped.jpg', { type: 'image/jpeg' });
    const preview = URL.createObjectURL(blob);
    setForm((prev) => ({ ...prev, image: file }));
    setImagePreview(preview);
    setCropperSrc(null);
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setSaving(true);
    try {
      const path = mode === 'create'
        ? '/api/offers/'
        : `/api/offers/${offer.id}/`;
      const method = mode === 'create' ? 'POST' : 'PUT';
      const hasImage = form.image instanceof File;

      let body;
      let headers = {};

      if (hasImage) {
        body = new FormData();
        body.append('title', form.title);
        body.append('description', form.description);
        body.append('image', form.image);
        body.append('start_date', form.start_date + ':00+05:30');
        body.append('end_date', form.end_date + ':00+05:30');
        body.append('terms_conditions', form.terms_conditions || '');
        body.append('is_active', form.is_active ? 'true' : 'false');
        selectedBranchIds.forEach((id) => body.append('branch_ids', id));
      } else {
        headers['Content-Type'] = 'application/json';
        body = JSON.stringify({
          title: form.title,
          description: form.description,
          start_date: form.start_date + ':00+05:30',
          end_date: form.end_date + ':00+05:30',
          terms_conditions: form.terms_conditions || '',
          is_active: form.is_active,
          branch_ids: selectedBranchIds,
        });
      }

      const res = await apiFetch(path, { method, headers, body });
      if (!res.ok) throw new Error('Failed to save offer');
      onSaved();
    } catch (err) {
      console.error(err);
      setError(err.message || 'Failed to save offer');
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card offer-modal-card" onClick={(e) => e.stopPropagation()} onSubmit={handleSubmit}>
        <div className="modal-head">
          <h3>{mode === 'create' ? 'Create Offer' : 'Edit Offer'}</h3>
          <button type="button" className="profile-close" onClick={onClose}><X size={18} /></button>
        </div>

        {error && <div className="login-error">{error}</div>}

        <div className="offer-form-grid">
          <div className="offer-form-left">
            <label className="modal-field">
              <span>Title</span>
              <input type="text" value={form.title} onChange={handleChange('title')} required />
            </label>

            <label className="modal-field">
              <span>Description</span>
              <textarea className="modal-textarea" value={form.description} onChange={handleChange('description')} rows={3} />
            </label>

            <label className="modal-field">
              <span>Image</span>
              <input type="file" accept="image/*" onChange={handleChange('image')} />
              {imagePreview && (
                <div className="offer-preview-wrap">
                  <img src={imagePreview} alt="Preview" className="offer-preview" />
                  <button type="button" className="offer-preview-remove" onClick={() => { setForm((prev) => ({ ...prev, image: null })); setImagePreview(null); }}>
                    <X size={14} />
                  </button>
                </div>
              )}
            </label>

            <label className="modal-field">
              <span>Terms & Conditions</span>
              <textarea className="modal-textarea" value={form.terms_conditions} onChange={handleChange('terms_conditions')} rows={3} />
            </label>
          </div>

          <div className="offer-form-right">
            <label className="modal-field">
              <span>Start Date</span>
              <input type="datetime-local" value={form.start_date} onChange={handleChange('start_date')} required />
            </label>

            <label className="modal-field">
              <span>End Date</span>
              <input type="datetime-local" value={form.end_date} onChange={handleChange('end_date')} required />
            </label>

            <label className="modal-field">
              <span>Branches</span>
              <BranchCheckDropdown selected={selectedBranchIds} onChange={setSelectedBranchIds} />
            </label>

            <label className="modal-check">
              <input type="checkbox" checked={form.is_active} onChange={handleChange('is_active')} />
              <span>Active</span>
            </label>
          </div>
        </div>

        <div className="modal-actions">
          <button type="button" className="small-button" onClick={onClose}>Cancel</button>
          <button type="submit" className="primary-button modal-submit" disabled={saving}>
            {saving ? 'Saving...' : mode === 'create' ? 'Create' : 'Update'}
          </button>
        </div>
      </form>
      {cropperSrc && (
        <ImageCropperModal
          imageSrc={cropperSrc}
          aspect={OFFER_IMAGE_ASPECT}
          onCrop={handleCropped}
          onCancel={() => setCropperSrc(null)}
        />
      )}
    </div>
  );
}

function OfferCalendar({ offers, onRefresh, onEditOffer }) {
  const today = new Date();
  const [weekStart, setWeekStart] = useState(() => {
    const d = new Date(today);
    d.setDate(d.getDate() - d.getDay() + 1);
    d.setHours(0, 0, 0, 0);
    return d;
  });

  const days = Array.from({ length: 7 }, (_, i) => {
    const d = new Date(weekStart);
    d.setDate(d.getDate() + i);
    return d;
  });

  const weekEnd = new Date(days[6]);
  weekEnd.setHours(23, 59, 59, 999);

  const dayWidth = 100 / 7;
  const now = new Date();

  const prevWeek = () => {
    const d = new Date(weekStart);
    d.setDate(d.getDate() - 7);
    setWeekStart(d);
  };

  const nextWeek = () => {
    const d = new Date(weekStart);
    d.setDate(d.getDate() + 7);
    setWeekStart(d);
  };

  const goToThisWeek = () => {
    const d = new Date();
    d.setDate(d.getDate() - d.getDay() + 1);
    d.setHours(0, 0, 0, 0);
    setWeekStart(d);
  };

  const getOfferBarStyle = (offer) => {
    const start = new Date(offer.start_date);
    const end = new Date(offer.end_date);
    const visibleStart = new Date(days[0]);
    visibleStart.setHours(0, 0, 0, 0);
    const visibleEnd = new Date(days[6]);
    visibleEnd.setHours(23, 59, 59, 999);

    const barStart = start < visibleStart ? visibleStart : start;
    const barEnd = end > visibleEnd ? visibleEnd : end;

    const totalMs = visibleEnd - visibleStart;
    const leftMs = barStart - visibleStart;
    const widthMs = barEnd - barStart;

    const left = (leftMs / totalMs) * 100;
    const width = Math.max((widthMs / totalMs) * 100, 1.5);

    return { left: `${left}%`, width: `${width}%` };
  };

  const getOfferColor = (offer) => {
    const start = new Date(offer.start_date);
    const end = new Date(offer.end_date);
    if (now > end) return 'cal-expired';
    if (now < start) return 'cal-scheduled';
    if (offer.is_active) return 'cal-active';
    return 'cal-inactive';
  };

  const [dragging, setDragging] = useState(null);

  const handleBarMouseDown = (e, offer) => {
    e.preventDefault();
    const trackEl = e.currentTarget.closest('.calendar-offers-track');
    const rect = trackEl.getBoundingClientRect();
    setDragging({
      mode: 'move',
      offer,
      startClientX: e.clientX,
      origStart: new Date(offer.start_date),
      origEnd: new Date(offer.end_date),
      rect,
      moved: false,
      dxPx: 0,
    });
  };

  const handleEdgeMouseDown = (e, offer, edge) => {
    e.stopPropagation();
    e.preventDefault();
    const trackEl = e.currentTarget.closest('.calendar-offers-track');
    const rect = trackEl.getBoundingClientRect();
    setDragging({
      mode: 'resize',
      offer,
      edge,
      startClientX: e.clientX,
      origStart: new Date(offer.start_date),
      origEnd: new Date(offer.end_date),
      rect,
      moved: false,
      dxPx: 0,
    });
  };

  useEffect(() => {
    if (!dragging) return;

    const handleMouseMove = (e) => {
      const dx = e.clientX - dragging.startClientX;
      if (Math.abs(dx) <= 2 && !dragging.moved) return;
      setDragging((prev) => prev ? { ...prev, moved: true, dxPx: dx } : prev);
    };

    const handleMouseUp = async (e) => {
      const dx = e.clientX - dragging.startClientX;

      if (!dragging.moved) {
        onEditOffer(dragging.offer);
        setDragging(null);
        return;
      }

      const totalMs = weekEnd - weekStart;
      const totalPx = dragging.rect.width;
      const msDelta = (dx / totalPx) * totalMs;

      let patch = {};

      if (dragging.mode === 'move') {
        const newStart = new Date(dragging.origStart.getTime() + msDelta);
        const newEnd = new Date(dragging.origEnd.getTime() + msDelta);
        newStart.setMinutes(0, 0, 0);
        newEnd.setMinutes(0, 0, 0);
        if (newStart.getHours() >= 12) newStart.setHours(12); else newStart.setHours(0);
        if (newEnd.getHours() >= 12) newEnd.setHours(12); else newEnd.setHours(0);
        if (newStart >= newEnd) { setDragging(null); return; }
        patch.start_date = newStart.toISOString().slice(0, 19) + '+05:30';
        patch.end_date = newEnd.toISOString().slice(0, 19) + '+05:30';
      } else if (dragging.mode === 'resize') {
        if (dragging.edge === 'left') {
          const newStart = new Date(dragging.origStart.getTime() + msDelta);
          newStart.setMinutes(0, 0, 0);
          if (newStart.getHours() >= 12) newStart.setHours(12); else newStart.setHours(0);
          if (newStart >= new Date(dragging.origEnd)) { setDragging(null); return; }
          patch.start_date = newStart.toISOString().slice(0, 19) + '+05:30';
        } else {
          const newEnd = new Date(dragging.origEnd.getTime() + msDelta);
          newEnd.setMinutes(0, 0, 0);
          if (newEnd.getHours() >= 12) newEnd.setHours(12); else newEnd.setHours(0);
          if (newEnd <= new Date(dragging.origStart)) { setDragging(null); return; }
          patch.end_date = newEnd.toISOString().slice(0, 19) + '+05:30';
        }
      }

      try {
        await apiFetch(`/api/offers/${dragging.offer.id}/`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(patch),
        });
        onRefresh();
      } catch (err) {
        console.error('Failed to update offer dates', err);
      }
      setDragging(null);
    };

    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
  }, [dragging, weekStart, weekEnd, onRefresh, onEditOffer]);

  const weekLabel = () => {
    const opts = { day: 'numeric', month: 'short', year: 'numeric' };
    return `${days[0].toLocaleDateString('en-IN', opts)} — ${days[6].toLocaleDateString('en-IN', opts)}`;
  };

  return (
    <div className="calendar-wrap">
      <div className="calendar-header">
        <div className="calendar-nav">
          <button className="calendar-nav-btn" onClick={prevWeek}><ChevronLeft size={18} /></button>
          <span className="calendar-week-label">{weekLabel()}</span>
          <button className="calendar-nav-btn" onClick={nextWeek}><ChevronRight size={18} /></button>
        </div>
        <button className="calendar-today-btn" onClick={goToThisWeek}>Today</button>
      </div>

      <div className="calendar-grid">
        <div className="calendar-day-headers">
          {days.map((d) => {
            const isToday = d.toDateString() === today.toDateString();
            return (
              <div key={d.toISOString()} className={`calendar-day-header ${isToday ? 'today' : ''}`}>
                <span className="cal-day-name">{d.toLocaleDateString('en-US', { weekday: 'short' })}</span>
                <span className="cal-day-num">{d.getDate()}</span>
              </div>
            );
          })}
        </div>

        <div className="calendar-timeline">
          {days.map((d) => {
            const isToday = d.toDateString() === today.toDateString();
            return <div key={d.toISOString()} className={`calendar-day-col ${isToday ? 'today' : ''}`} />;
          })}

          <div className="calendar-offers-track">
            {offers.map((o) => {
              const style = getOfferBarStyle(o);
              const colorClass = getOfferColor(o);
              const status = now > new Date(o.end_date) ? 'Expired' : now < new Date(o.start_date) ? 'Scheduled' : o.is_active ? 'Active' : 'Inactive';
              const isDragging = dragging && dragging.offer.id === o.id && dragging.moved;
              const isResizing = isDragging && dragging.mode === 'resize';
              const isMoving = isDragging && dragging.mode === 'move';
              let barStyle = { ...style };
              if (isMoving) {
                barStyle = { ...barStyle, transform: `translateX(${dragging.dxPx}px)`, transition: 'none', zIndex: 10 };
              } else if (isResizing) {
                const totalMs = weekEnd - weekStart;
                const msDelta = (dragging.dxPx / dragging.rect.width) * totalMs;
                if (dragging.edge === 'left') {
                  const origLeftMs = new Date(dragging.origStart) - weekStart;
                  const newLeftMs = origLeftMs + msDelta;
                  const newLeftPct = (newLeftMs / (weekEnd - weekStart)) * 100;
                  const origWidthMs = new Date(dragging.origEnd) - new Date(dragging.origStart);
                  barStyle.left = `${Math.max(0, newLeftPct)}%`;
                  barStyle.width = style.width;
                  barStyle.transition = 'none';
                  barStyle.zIndex = 10;
                } else {
                  const origWidthMs = new Date(dragging.origEnd) - new Date(dragging.origStart);
                  const newWidthMs = origWidthMs + msDelta;
                  const newWidthPct = (newWidthMs / (weekEnd - weekStart)) * 100;
                  barStyle.width = `${Math.max(2, newWidthPct)}%`;
                  barStyle.transition = 'none';
                  barStyle.zIndex = 10;
                }
              }
              return (
                <div
                  key={o.id}
                  className={`cal-offer-bar ${colorClass} ${isDragging ? 'dragging' : ''}`}
                  style={barStyle}
                  onMouseDown={(e) => handleBarMouseDown(e, o)}
                  title={`${o.title} (${status})\nDrag to move, click to edit`}
                >
                  <div className="cal-edge-left" onMouseDown={(e) => handleEdgeMouseDown(e, o, 'left')} />
                  <span className="cal-offer-label">{o.title}</span>
                  <div className="cal-edge-right" onMouseDown={(e) => handleEdgeMouseDown(e, o, 'right')} />
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}

function OffersPage() {
  const [offers, setOffers] = useState([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(1);
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const [loading, setLoading] = useState(true);
  const [modal, setModal] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);
  const [selectedIds, setSelectedIds] = useState([]);
  const [bulkBusy, setBulkBusy] = useState(false);
  const [viewMode, setViewMode] = useState('table');
  const pageSize = 10;

  const fetchOffers = useCallback(async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, page_size: pageSize });
      if (search) params.set('search', search);
      const res = await apiFetch(`/api/offers/admin_all/?${params}`);
      const data = await res.json();
      setOffers(Array.isArray(data) ? data : data.results || []);
      setTotal(Array.isArray(data) ? data.length : data.count || 0);
    } catch (err) {
      console.error('Failed to fetch offers', err);
    } finally {
      setLoading(false);
    }
  }, [page, search]);

  useEffect(() => { fetchOffers(); }, [fetchOffers]);

  const now = new Date();

  const getOfferStatus = (offer) => {
    const start = new Date(offer.start_date);
    const end = new Date(offer.end_date);
    if (now > end) return 'Expired';
    if (now < start) return 'Scheduled';
    if (offer.is_active) return 'Active';
    return 'Inactive';
  };

  const filteredOffers = offers.filter((o) => {
    const status = getOfferStatus(o);
    if (statusFilter === 'active') return status === 'Active';
    if (statusFilter === 'scheduled') return status === 'Scheduled';
    if (statusFilter === 'inactive') return status === 'Inactive';
    if (statusFilter === 'expired') return status === 'Expired';
    return true;
  });

  const handleDelete = async () => {
    if (!deleteTarget) return;
    try {
      await apiFetch(`/api/offers/${deleteTarget.id}/`, {
        method: 'DELETE',
      });
      setDeleteTarget(null);
      fetchOffers();
    } catch (err) {
      console.error('Failed to delete offer', err);
    }
  };

  const toggleSelect = (id) => {
    setSelectedIds((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  };

  const selectAll = () => {
    const allIds = filteredOffers.map((o) => o.id);
    setSelectedIds((prev) => prev.length === allIds.length ? [] : allIds);
  };

  const handleBulkDelete = async () => {
    if (!selectedIds.length) return;
    setBulkBusy(true);
    await Promise.allSettled(
      selectedIds.map((id) =>
        apiFetch(`/api/offers/${id}/`, {
          method: 'DELETE',
        })
      )
    );
    setSelectedIds([]);
    setBulkBusy(false);
    fetchOffers();
  };

  const handleBulkSetActive = async () => {
    if (!selectedIds.length) return;
    setBulkBusy(true);
    await Promise.allSettled(
      selectedIds.map((id) =>
        apiFetch(`/api/offers/${id}/`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ is_active: true }),
        })
      )
    );
    setSelectedIds([]);
    setBulkBusy(false);
    setStatusFilter('all');
    setPage(1);
    fetchOffers();
  };

  const handleBulkSetInactive = async () => {
    if (!selectedIds.length) return;
    setBulkBusy(true);
    await Promise.allSettled(
      selectedIds.map((id) =>
        apiFetch(`/api/offers/${id}/`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ is_active: false }),
        })
      )
    );
    setSelectedIds([]);
    setBulkBusy(false);
    setStatusFilter('all');
    setPage(1);
    fetchOffers();
  };

  const totalPages = Math.ceil(total / pageSize);

  const filterTabs = [
    { key: 'all', label: 'All' },
    { key: 'active', label: 'Active' },
    { key: 'scheduled', label: 'Scheduled' },
    { key: 'inactive', label: 'Inactive' },
    { key: 'expired', label: 'Expired' },
  ];

  const formatDate = (iso) => {
    if (!iso) return '—';
    return new Date(iso).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
  };
  const exportColumns = isExisting
    ? [
        { header: 'Code', getValue: (r) => r.code || '—' },
        { header: 'Name', getValue: (r) => r.name },
        { header: 'Phone', getValue: (r) => r.phone2 || '—' },
        { header: 'Place', getValue: (r) => r.place || '—' },
        { header: 'Super Code', getValue: (r) => r.super_code || '—' },
        { header: 'Client ID', getValue: (r) => r.client_id || '—' },
        { header: 'Synced', getValue: (r) => formatDate(r.synced_at) },
      ]
    : [
        { header: 'Name', getValue: (r) => r.name },
        { header: 'Phone', getValue: (r) => r.phone },
        { header: 'Email', getValue: (r) => r.email || '—' },
        { header: 'Role', getValue: (r) => roleLabel(r.role) },
        { header: 'Joined', getValue: (r) => formatDate(r.date_joined) },
        { header: 'Status', getValue: (r) => r.is_active ? 'Active' : 'Inactive' },
      ];

  return (
    <>
      <div className="hero-row">
        <div>
          <p>Manage Offers</p>
          <h2>Offers</h2>
        </div>
        <div className="filters">
          <div className="search-field">
            <Search size={17} />
            <input type="text" placeholder="Search offers..." value={search} onChange={(e) => { setSearch(e.target.value); setPage(1); }} />
          </div>
          <button className="primary-button" onClick={() => setModal({ mode: 'create' })}>
            <Plus size={18} /> Add Offer
          </button>
          <button className="export-button" onClick={() => exportToExcel(filteredOffers, [
            { header: 'Title', getValue: (r) => r.title },
            { header: 'Description', getValue: (r) => r.description },
            { header: 'Start Date', getValue: (r) => formatDate(r.start_date) },
            { header: 'End Date', getValue: (r) => formatDate(r.end_date) },
            { header: 'Status', getValue: (r) => getOfferStatus(r) },
            { header: 'Views', getValue: (r) => r.view_count },
          ], 'offers')}>
            <Download size={16} /> Export
          </button>
          <div className="view-toggle">
            <button className={`view-toggle-btn ${viewMode === 'table' ? 'active' : ''}`} onClick={() => setViewMode('table')} title="Table view"><List size={16} /></button>
            <button className={`view-toggle-btn ${viewMode === 'calendar' ? 'active' : ''}`} onClick={() => setViewMode('calendar')} title="Calendar view"><CalendarDays size={16} /></button>
          </div>
        </div>
      </div>

      {viewMode === 'calendar' ? (
        <section className="panel offers-panel">
          <div className="filter-tabs">
            {filterTabs.map((t) => (
              <button key={t.key} className={`filter-tab ${statusFilter === t.key ? 'active' : ''}`} onClick={() => { setStatusFilter(t.key); setPage(1); }}>
                {t.label}
              </button>
            ))}
          </div>
          {loading ? (
            <p className="table-loading">Loading offers...</p>
          ) : (
            <OfferCalendar offers={filteredOffers} onRefresh={fetchOffers} onEditOffer={(o) => setModal({ mode: 'edit', offer: o })} />
          )}
        </section>
      ) : (
      <section className="panel offers-panel">
        <div className="filter-tabs">
          {filterTabs.map((t) => (
            <button key={t.key} className={`filter-tab ${statusFilter === t.key ? 'active' : ''}`} onClick={() => { setStatusFilter(t.key); setPage(1); }}>
              {t.label}
            </button>
          ))}
        </div>

        {loading ? (
          <p className="table-loading">Loading offers...</p>
        ) : filteredOffers.length === 0 ? (
          search || statusFilter !== 'all' ? (
            <EmptyState icon={Search} title="No matching offers" description="No offers match your current filters. Try adjusting your search or status filter." />
          ) : (
            <EmptyState icon={Gift} title="No offers yet" description="Create your first offer to start engaging customers." actionLabel="Create Offer" onAction={() => setModal({ mode: 'create' })} />
          )
        ) : (
          <div className="offers-table">
            <div className="offers-table-head">
              <label className="bulk-checkbox-cell">
                <input type="checkbox" checked={selectedIds.length === filteredOffers.length && filteredOffers.length > 0} onChange={selectAll} />
              </label>
              <span>Title</span>
              <span>Period</span>
              <span>Status</span>
              <span>Views</span>
              <span>Actions</span>
            </div>
            {filteredOffers.map((o) => (
              <div className="offers-table-row" key={o.id}>
                <label className="bulk-checkbox-cell">
                  <input type="checkbox" checked={selectedIds.includes(o.id)} onChange={() => toggleSelect(o.id)} />
                </label>
                <div className="offer-title-cell">
                  {o.image && <img src={mediaUrl(o.image)} alt="" className="offer-thumb" />}
                  <div>
                    <strong>{o.title}</strong>
                    <span>{o.description}</span>
                  </div>
                </div>
                <span className="offer-period">{formatDate(o.start_date)} - {formatDate(o.end_date)}</span>
                <span className="offer-status-cell">
                  <span className={`status ${getOfferStatus(o) === 'Active' ? '' : getOfferStatus(o) === 'Scheduled' ? 'scheduled' : 'pending'}`}>{getOfferStatus(o)}</span>
                </span>
                <span className="offer-views">{o.view_count}</span>
                <span className="offer-actions">
                  <button onClick={() => setModal({ mode: 'edit', offer: o })} aria-label="Edit"><Edit3 size={16} /></button>
                  <button onClick={() => setDeleteTarget(o)} aria-label="Delete"><Trash2 size={16} /></button>
                </span>
              </div>
            ))}
          </div>
        )}

        {totalPages > 1 && (
          <div className="pagination">
            <button disabled={page <= 1} onClick={() => setPage((p) => p - 1)}><ChevronLeft size={16} /> Prev</button>
            <span>Page {page} of {totalPages}</span>
            <button disabled={page >= totalPages} onClick={() => setPage((p) => p + 1)}>Next <ChevronRight size={16} /></button>
          </div>
        )}
      </section>
      )}

      {modal && (
        <OfferFormModal
          mode={modal.mode}
          offer={modal.offer}
          onClose={() => setModal(null)}
          onSaved={() => { setModal(null); fetchOffers(); }}
        />
      )}

      {deleteTarget && (
        <div className="modal-overlay" onClick={() => setDeleteTarget(null)}>
          <div className="modal-card confirm-card" onClick={(e) => e.stopPropagation()}>
            <h3>Delete Offer</h3>
            <p>Are you sure you want to delete <strong>{deleteTarget.title}</strong>?</p>
            <div className="modal-actions">
              <button className="small-button" onClick={() => setDeleteTarget(null)}>Cancel</button>
              <button className="danger-button" onClick={handleDelete}>Delete</button>
            </div>
          </div>
        </div>
      )}

      {selectedIds.length > 0 && (
        <BulkActionBar
          count={selectedIds.length}
          onDelete={handleBulkDelete}
          onSetActive={handleBulkSetActive}
          onSetInactive={handleBulkSetInactive}
          onClear={() => setSelectedIds([])}
          busy={bulkBusy}
        />
      )}
    </>
  );
}

function StaffFormModal({ mode, staff, onClose, onSaved }) {
  const [form, setForm] = useState({
    name: staff?.name || '',
    phone: staff?.phone || '',
    password: '',
    role: staff?.role || 'BRANCH_MGR',
    managed_branch: staff?.managed_branch || '',
    is_active: staff?.is_active ?? true,
  });
  const [branches, setBranches] = useState([]);
  const [saving, setSaving] = useState(false);
  const [changePassword, setChangePassword] = useState(false);
  const [confirmPassword, setConfirmPassword] = useState('');
  const [error, setError] = useState('');

  useEffect(() => {
    apiFetch(`/api/branches/`)
      .then((r) => r.json())
      .then((data) => setBranches(data.results || []))
      .catch(() => {});
  }, []);

  const handleChange = (field) => (e) => {
    const value = field === 'is_active' ? e.target.checked : e.target.value;
    setForm((prev) => ({ ...prev, [field]: value }));
  };

  const cancelPasswordChange = () => {
    setChangePassword(false);
    setConfirmPassword('');
    setError('');
    setForm((prev) => ({ ...prev, password: '' }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError('');
    const isCreate = mode === 'create';

    if (isCreate) {
      if (form.password.length < 8) {
        setError('Password must be at least 8 characters.');
        return;
      }
    } else if (changePassword) {
      if (form.password.length < 8) {
        setError('Password must be at least 8 characters.');
        return;
      }
      if (form.password !== confirmPassword) {
        setError('Passwords do not match.');
        return;
      }
    }

    setSaving(true);
    try {
      const path = isCreate ? '/api/staff/' : `/api/staff/${staff.id}/`;
      const method = isCreate ? 'POST' : 'PATCH';

      const body = {
        name: form.name,
        phone: form.phone,
        role: form.role,
        is_active: form.is_active,
      };
      if (isCreate || changePassword) body.password = form.password;
      if (form.managed_branch) body.managed_branch = parseInt(form.managed_branch);

      const res = await apiFetch(path, {
        method,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });

      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        const firstError = Object.values(data)[0];
        throw new Error(
          Array.isArray(firstError) ? firstError[0] : firstError || 'Failed to save staff'
        );
      }
      onSaved();
    } catch (err) {
      setError(err.message || 'Failed to save staff');
      console.error(err);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card staff-modal-card" onClick={(e) => e.stopPropagation()} onSubmit={handleSubmit}>
        <div className="modal-head">
          <h3>{mode === 'create' ? 'Add Staff' : 'Edit Staff'}</h3>
          <button type="button" className="profile-close" onClick={onClose}><X size={18} /></button>
        </div>

        <div className="staff-form-grid">
          <div className="modal-field">
            <span>Name</span>
            <input type="text" value={form.name} onChange={handleChange('name')} required />
          </div>

          <div className="modal-field">
            <span>Phone</span>
            <input type="tel" inputMode="numeric" pattern="[0-9]{10}" minLength="10" maxLength="10" title="Enter a valid 10-digit phone number" value={form.phone} onChange={handleChange('phone')} required />
          </div>

          {mode === 'create' ? (
            <div className="modal-field">
              <span>Password</span>
              <input type="password" value={form.password} onChange={handleChange('password')} minLength="8" autoComplete="new-password" required />
            </div>
          ) : (
            <div className="modal-field">
              <span>Password</span>
              {changePassword ? (
                <div className="password-change-fields">
                  <input type="password" placeholder="New password" value={form.password} onChange={handleChange('password')} autoComplete="new-password" />
                  <input type="password" placeholder="Confirm new password" value={confirmPassword} onChange={(e) => setConfirmPassword(e.target.value)} autoComplete="new-password" />
                  <button type="button" className="small-button" onClick={cancelPasswordChange}>Cancel</button>
                </div>
              ) : (
                <div className="password-change-row">
                  <span className="password-masked">••••••••</span>
                  <button
                    type="button"
                    className="small-button"
                    onClick={() => { setChangePassword(true); setError(''); }}
                  >
                    Change password
                  </button>
                </div>
              )}
            </div>
          )}

          <div className="modal-field">
            <span>Role</span>
            <select className="modal-select" value={form.role} onChange={handleChange('role')}>
              <option value="BRANCH_MGR">Branch Manager</option>
              <option value="MARKETING">Marketing Manager</option>
            </select>
          </div>

          {form.role === 'BRANCH_MGR' && (
            <div className="modal-field">
              <span>Managed Branch</span>
              <select className="modal-select" value={form.managed_branch} onChange={handleChange('managed_branch')}>
                <option value="">Select branch...</option>
                {branches.map((b) => (
                  <option key={b.id} value={b.id}>{b.name}</option>
                ))}
              </select>
            </div>
          )}

          <div className={`modal-field${form.role === 'BRANCH_MGR' ? '' : ' staff-field-wide'}`}>
            <span>Status</span>
            <label className="modal-check">
              <input type="checkbox" checked={form.is_active} onChange={handleChange('is_active')} />
              <span>Active</span>
            </label>
          </div>
        </div>

        {error && <div className="modal-error">{error}</div>}

        <div className="modal-actions">
          <button type="button" className="small-button" onClick={onClose}>Cancel</button>
          <button type="submit" className="primary-button modal-submit" disabled={saving}>
            {saving ? 'Saving...' : mode === 'create' ? 'Create' : 'Update'}
          </button>
        </div>
      </form>
    </div>
  );
}

function UsersPage() {
  const [staff, setStaff] = useState([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(1);
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const [loading, setLoading] = useState(true);
  const [modal, setModal] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);
  const [selectedIds, setSelectedIds] = useState([]);
  const [bulkBusy, setBulkBusy] = useState(false);
  const pageSize = 10;

  const fetchStaff = useCallback(async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, page_size: pageSize });
      if (search) params.set('search', search);
      const res = await apiFetch(`/api/staff/?${params}`);
      const data = await res.json();
      setStaff(data.results || []);
      setTotal(data.count || 0);
    } catch (err) {
      console.error('Failed to fetch staff', err);
    } finally {
      setLoading(false);
    }
  }, [page, search]);

  useEffect(() => { fetchStaff(); }, [fetchStaff]);

  const filteredStaff = staff.filter((s) => {
    if (statusFilter === 'active') return s.is_active === true;
    if (statusFilter === 'inactive') return s.is_active === false;
    return true;
  });

  const handleDelete = async () => {
    if (!deleteTarget) return;
    try {
      await apiFetch(`/api/staff/${deleteTarget.id}/`, {
        method: 'DELETE',
      });
      setDeleteTarget(null);
      fetchStaff();
    } catch (err) {
      console.error('Failed to delete staff', err);
    }
  };

  const toggleSelect = (id) => {
    setSelectedIds((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  };

  const selectAll = () => {
    const allIds = filteredStaff.map((s) => s.id);
    setSelectedIds((prev) => prev.length === allIds.length ? [] : allIds);
  };

  const handleBulkDelete = async () => {
    if (!selectedIds.length) return;
    setBulkBusy(true);
    await Promise.allSettled(
      selectedIds.map((id) =>
        apiFetch(`/api/staff/${id}/`, {
          method: 'DELETE',
        })
      )
    );
    setSelectedIds([]);
    setBulkBusy(false);
    fetchStaff();
  };

  const handleBulkSetActive = async () => {
    if (!selectedIds.length) return;
    setBulkBusy(true);
    await Promise.allSettled(
      selectedIds.map((id) =>
        apiFetch(`/api/staff/${id}/`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ is_active: true }),
        })
      )
    );
    setSelectedIds([]);
    setBulkBusy(false);
    setStatusFilter('all');
    setPage(1);
    fetchStaff();
  };

  const handleBulkSetInactive = async () => {
    if (!selectedIds.length) return;
    setBulkBusy(true);
    await Promise.allSettled(
      selectedIds.map((id) =>
        apiFetch(`/api/staff/${id}/`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ is_active: false }),
        })
      )
    );
    setSelectedIds([]);
    setBulkBusy(false);
    setStatusFilter('all');
    setPage(1);
    fetchStaff();
  };

  const totalPages = Math.ceil(total / pageSize);

  const filterTabs = [
    { key: 'all', label: 'All' },
    { key: 'active', label: 'Active' },
    { key: 'inactive', label: 'Inactive' },
  ];

  const roleLabel = (role) => {
    if (role === 'BRANCH_MGR') return 'Branch Manager';
    if (role === 'MARKETING') return 'Marketing';
    return role;
  };

  const formatDate = (iso) => {
    if (!iso) return '—';
    return new Date(iso).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
  };

  return (
    <>
      <div className="hero-row">
        <div>
          <p>Manage Users</p>
          <h2>Staff</h2>
        </div>
        <div className="filters">
          <div className="search-field">
            <Search size={17} />
            <input type="text" placeholder="Search staff..." value={search} onChange={(e) => { setSearch(e.target.value); setPage(1); }} />
          </div>
          <button className="primary-button" onClick={() => setModal({ mode: 'create' })}>
            <Plus size={18} /> Add Staff
          </button>
          <button className="export-button" onClick={() => exportToExcel(filteredStaff, [
            { header: 'Name', getValue: (r) => r.name },
            { header: 'Phone', getValue: (r) => r.phone },
            { header: 'Role', getValue: (r) => r.role === 'BRANCH_MGR' ? 'Branch Manager' : 'Marketing' },
            { header: 'Branch', getValue: (r) => r.managed_branch_name || '—' },
            { header: 'Joined', getValue: (r) => r.date_joined ? new Date(r.date_joined).toLocaleDateString('en-IN') : '—' },
            { header: 'Status', getValue: (r) => r.is_active ? 'Active' : 'Inactive' },
          ], 'staff')}>
            <Download size={16} /> Export
          </button>
        </div>
      </div>

      <section className="panel staff-panel">
        <div className="filter-tabs">
          {filterTabs.map((t) => (
            <button key={t.key} className={`filter-tab ${statusFilter === t.key ? 'active' : ''}`} onClick={() => { setStatusFilter(t.key); setPage(1); }}>
              {t.label}
            </button>
          ))}
        </div>

        {loading ? (
          <p className="table-loading">Loading staff...</p>
        ) : filteredStaff.length === 0 ? (
          search || statusFilter !== 'all' ? (
            <EmptyState icon={Search} title="No matching staff" description="No staff members match your current filters. Try adjusting your search." />
          ) : (
            <EmptyState icon={Users} title="No staff yet" description="Add your first staff member to start managing your team." actionLabel="Add Staff" onAction={() => setModal({ mode: 'create' })} />
          )
        ) : (
          <div className="staff-table">
            <div className="staff-table-head">
              <label className="bulk-checkbox-cell">
                <input type="checkbox" checked={selectedIds.length === filteredStaff.length && filteredStaff.length > 0} onChange={selectAll} />
              </label>
              <span>Staff</span>
              <span>Role</span>
              <span>Branch</span>
              <span>Joined</span>
              <span>Status</span>
              <span>Actions</span>
            </div>
            {filteredStaff.map((s) => (
              <div className="staff-table-row" key={s.id}>
                <label className="bulk-checkbox-cell">
                  <input type="checkbox" checked={selectedIds.includes(s.id)} onChange={() => toggleSelect(s.id)} />
                </label>
                <div className="staff-info">
                  <span className="staff-avatar">{s.name?.[0] || 'S'}</span>
                  <div>
                    <strong>{s.name}</strong>
                    <span>{s.phone}</span>
                  </div>
                </div>
                <span className={`staff-role ${s.role === 'BRANCH_MGR' ? 'role-mgr' : 'role-mkt'}`}>
                  {s.role === 'BRANCH_MGR' ? <Store size={13} /> : <ShieldCheck size={13} />}
                  {roleLabel(s.role)}
                </span>
                <span className="staff-branch">{s.managed_branch_name || '—'}</span>
                <span className="staff-joined">{formatDate(s.date_joined)}</span>
                <span className="staff-status-cell">
                  <span className={`status ${!s.is_active ? 'pending' : ''}`}>{s.is_active ? 'Active' : 'Inactive'}</span>
                </span>
                <span className="staff-actions">
                  <button onClick={() => setModal({ mode: 'edit', staff: s })} aria-label="Edit"><Edit3 size={16} /></button>
                  <button onClick={() => setDeleteTarget(s)} aria-label="Delete"><Trash2 size={16} /></button>
                </span>
              </div>
            ))}
          </div>
        )}

        {totalPages > 1 && (
          <div className="pagination">
            <button disabled={page <= 1} onClick={() => setPage((p) => p - 1)}><ChevronLeft size={16} /> Prev</button>
            <span>Page {page} of {totalPages}</span>
            <button disabled={page >= totalPages} onClick={() => setPage((p) => p + 1)}>Next <ChevronRight size={16} /></button>
          </div>
        )}
      </section>

      {modal && (
        <StaffFormModal
          mode={modal.mode}
          staff={modal.staff}
          onClose={() => setModal(null)}
          onSaved={() => { setModal(null); fetchStaff(); }}
        />
      )}

      {deleteTarget && (
        <div className="modal-overlay" onClick={() => setDeleteTarget(null)}>
          <div className="modal-card confirm-card" onClick={(e) => e.stopPropagation()}>
            <h3>Delete Staff</h3>
            <p>Are you sure you want to delete <strong>{deleteTarget.name}</strong>?</p>
            <div className="modal-actions">
              <button className="small-button" onClick={() => setDeleteTarget(null)}>Cancel</button>
              <button className="danger-button" onClick={handleDelete}>Delete</button>
            </div>
          </div>
        </div>
      )}

      {selectedIds.length > 0 && (
        <BulkActionBar
          count={selectedIds.length}
          onDelete={handleBulkDelete}
          onSetActive={handleBulkSetActive}
          onSetInactive={handleBulkSetInactive}
          onClear={() => setSelectedIds([])}
          busy={bulkBusy}
        />
      )}
    </>
  );
}

function AccountUsersPage() {
  const [source, setSource] = useState('app');
  const [rows, setRows] = useState([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(1);
  const [search, setSearch] = useState('');
  const [role, setRole] = useState('');
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const pageSize = 20;
  const isExisting = source === 'existing';

  const fetchRows = useCallback(async () => {
    setLoading(true);
    setError('');
    try {
      const params = new URLSearchParams({ page: String(page) });
      if (search) params.set('search', search);
      if (!isExisting && role) params.set('role', role);
      const url = isExisting
        ? `${ACC_MASTER_API_URL}?${params}`
        : `/api/users/?${params}`;
      const res = await apiFetch(url);
      if (!res.ok) throw new Error(`Request failed (${res.status})`);
      const data = await res.json();
      setRows(data.results || []);
      setTotal(data.count || 0);
    } catch (err) {
      console.error('Failed to fetch customers', err);
      setError(isExisting ? 'Could not load existing customers.' : 'Could not load users from accounts_user.');
    } finally {
      setLoading(false);
    }
  }, [page, search, role, isExisting]);

  useEffect(() => { fetchRows(); }, [fetchRows]);

  const changeSource = (next) => {
    if (next === source) return;
    setSource(next);
    setPage(1);
    setSearch('');
    setRole('');
  };

  const totalPages = Math.ceil(total / pageSize);
  const roleLabel = (value) => {
    if (value === 'BRANCH_MGR') return 'Branch Manager';
    if (value === 'SUPER_ADMIN') return 'Super Admin';
    if (value === 'MARKETING') return 'Marketing';
    return value;
  };
  const formatDate = (iso) => {
    if (!iso) return '—';
    return new Date(iso).toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
  };

  return (
    <>
      <div className="hero-row">
        <div>
          <p>Account users</p>
          <h2>Customers</h2>
        </div>
        <div className="filters">
          <div className="search-field">
            <Search size={17} />
            <input
              type="text"
              placeholder={isExisting ? 'Search code, name, place...' : 'Search phone, name, email...'}
              value={search}
              onChange={(e) => { setSearch(e.target.value); setPage(1); }}
            />
          </div>
          {!isExisting && (
            <div className="search-field">
              <Filter size={17} />
              <select value={role} onChange={(e) => { setRole(e.target.value); setPage(1); }} aria-label="Filter by role">
                <option value="">All roles</option>
                <option value="CUSTOMER">Customer</option>
                <option value="SUPER_ADMIN">Super Admin</option>
                <option value="BRANCH_MGR">Branch Manager</option>
                <option value="MARKETING">Marketing</option>
              </select>
            </div>
          )}
          <button className="export-button" onClick={() => exportToExcel(rows, exportColumns, isExisting ? 'existing-customers' : 'users')}>
            <Download size={16} /> Export
          </button>
        </div>
      </div>

      <div className="filter-tabs customer-source-tabs">
        <button type="button" className={`filter-tab ${!isExisting ? 'active' : ''}`} onClick={() => changeSource('app')}>
          App Registered
        </button>
        <button type="button" className={`filter-tab ${isExisting ? 'active' : ''}`} onClick={() => changeSource('existing')}>
          Existing Customers
        </button>
      </div>

      <section className="panel">
        {error && <div className="login-error">{error}</div>}
        {loading ? (
          <p className="table-loading">Loading customers...</p>
        ) : rows.length === 0 ? (
          <EmptyState
            icon={Users}
            title="No customers found"
            description={isExisting ? 'No acc-master records match the current filters.' : 'No accounts_user records match the current filters.'}
          />
        ) : (
          <div className="customer-table">
            <div className="customer-table-head">
              {isExisting ? (
                <>
                  <span>Code</span>
                  <span>Customer</span>
                  <span>Phone</span>
                  <span className="customer-place">Place</span>
                  <span className="customer-client">Super Code</span>
                  <span>Synced</span>
                </>
              ) : (
                <>
                  <span>User</span>
                  <span>Phone</span>
                  <span className="customer-place">Role</span>
                  <span className="customer-client">Email</span>
                  <span>Joined</span>
                  <span>Status</span>
                </>
              )}
            </div>
            {rows.map((row) => (
              <div className="customer-table-row" key={row.id}>
                {isExisting ? (
                  <>
                    <span>{row.code || '—'}</span>
                    <div className="customer-info">
                      <span className="staff-avatar">{row.name?.[0] || 'C'}</span>
                      <div>
                        <strong>{row.name}</strong>
                        <span>{row.client_id || '—'}</span>
                      </div>
                    </div>
                    <span>{row.phone2 || '—'}</span>
                    <span className="customer-place">{row.place || '—'}</span>
                    <span className="customer-client">{row.super_code || '—'}</span>
                    <span>{formatDate(row.synced_at)}</span>
                  </>
                ) : (
                  <>
                    <div className="customer-info">
                      <span className="staff-avatar">{row.name?.[0] || 'U'}</span>
                      <div>
                        <strong>{row.name}</strong>
                        <span>{row.first_name || row.last_name ? `${row.first_name || ''} ${row.last_name || ''}`.trim() : row.role}</span>
                      </div>
                    </div>
                    <span>{row.phone}</span>
                    <span className={`customer-place staff-role ${row.role === 'CUSTOMER' ? 'role-mkt' : 'role-mgr'}`}>
                      {roleLabel(row.role)}
                    </span>
                    <span className="customer-client">{row.email || '—'}</span>
                    <span>{formatDate(row.date_joined)}</span>
                    <span className={`status ${!row.is_active ? 'pending' : ''}`}>{row.is_active ? 'Active' : 'Inactive'}</span>
                  </>
                )}
              </div>
            ))}
          </div>
        )}

        {totalPages > 1 && (
          <div className="pagination">
            <button disabled={page <= 1} onClick={() => setPage((p) => p - 1)}><ChevronLeft size={16} /> Prev</button>
            <span>Page {page} of {totalPages}</span>
            <button disabled={page >= totalPages} onClick={() => setPage((p) => p + 1)}>Next <ChevronRight size={16} /></button>
          </div>
        )}
      </section>
    </>
  );
}

function BannerFormModal({ mode, banner, onClose, onSaved }) {
  const [form, setForm] = useState({
    title: banner?.title || '',
    image: null,
    link_url: banner?.link_url || '',
    linked_offer: banner?.linked_offer?.toString() || '',
    order: banner?.order?.toString() || '',
    start_date: banner?.start_date ? banner.start_date.slice(0, 16) : '',
    end_date: banner?.end_date ? banner.end_date.slice(0, 16) : '',
    is_active: banner?.is_active ?? true,
  });
  const [selectedBranchIds, setSelectedBranchIds] = useState(banner?.branches || []);
  const [imagePreview, setImagePreview] = useState(banner?.image ? mediaUrl(banner.image) : null);
  const [offers, setOffers] = useState([]);
  const [saving, setSaving] = useState(false);
  const [cropperSrc, setCropperSrc] = useState(null);

  useEffect(() => {
    apiFetch(`/api/offers/`)
      .then((r) => r.json())
      .then((data) => setOffers(data.results || []))
      .catch(() => {});
  }, []);

  const handleChange = (field) => (e) => {
    if (field === 'image') {
      const file = e.target.files[0];
      if (file) {
        const preview = URL.createObjectURL(file);
        setCropperSrc(preview);
      }
    } else if (field === 'is_active') {
      setForm((prev) => ({ ...prev, is_active: e.target.checked }));
    } else {
      setForm((prev) => ({ ...prev, [field]: e.target.value }));
    }
  };

  const handleCropped = (blob) => {
    const file = new File([blob], 'cropped.jpg', { type: 'image/jpeg' });
    const preview = URL.createObjectURL(blob);
    setForm((prev) => ({ ...prev, image: file }));
    setImagePreview(preview);
    setCropperSrc(null);
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setSaving(true);
    try {
      const path = mode === 'create'
        ? '/api/banners/'
        : `/api/banners/${banner.id}/`;
      const method = mode === 'create' ? 'POST' : 'PUT';
      const hasImage = form.image instanceof File;

      let body;
      let headers = {};

      if (hasImage) {
        body = new FormData();
        body.append('title', form.title);
        body.append('image', form.image);
        if (form.link_url) body.append('link_url', form.link_url);
        if (form.linked_offer) body.append('linked_offer', form.linked_offer);
        if (form.order) body.append('order', form.order);
        if (form.start_date) body.append('start_date', form.start_date + ':00+05:30');
        if (form.end_date) body.append('end_date', form.end_date + ':00+05:30');
        body.append('is_active', form.is_active ? 'true' : 'false');
        selectedBranchIds.forEach((id) => body.append('branches', id));
      } else {
        headers['Content-Type'] = 'application/json';
        body = JSON.stringify({
          title: form.title,
          link_url: form.link_url || '',
          linked_offer: form.linked_offer ? parseInt(form.linked_offer) : null,
          order: form.order ? parseInt(form.order) : null,
          start_date: form.start_date ? form.start_date + ':00+05:30' : null,
          end_date: form.end_date ? form.end_date + ':00+05:30' : null,
          is_active: form.is_active,
          branches: selectedBranchIds,
        });
      }

      const res = await apiFetch(path, { method, headers, body });
      if (!res.ok) {
        let errorText = 'Failed to save banner';
        try { const errData = await res.json(); errorText = JSON.stringify(errData); } catch {}
        throw new Error(errorText);
      }
      onSaved();
    } catch (err) {
      console.error('Banner save error:', err);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card banner-modal-card" onClick={(e) => e.stopPropagation()} onSubmit={handleSubmit}>
        <div className="modal-head">
          <h3>{mode === 'create' ? 'Add Banner' : 'Edit Banner'}</h3>
          <button type="button" className="profile-close" onClick={onClose}><X size={18} /></button>
        </div>

        <div className="banner-form-grid">
          <div className="banner-form-left">
            <label className="modal-field">
              <span>Title</span>
              <input type="text" value={form.title} onChange={handleChange('title')} required />
            </label>

            <label className="modal-field">
              <span>Image</span>
              <input type="file" accept="image/*" onChange={handleChange('image')} />
              {imagePreview && (
                <div className="offer-preview-wrap">
                  <img src={imagePreview} alt="Preview" className="banner-preview" />
                  <button type="button" className="offer-preview-remove" onClick={() => { setForm((prev) => ({ ...prev, image: null })); setImagePreview(null); }}>
                    <X size={14} />
                  </button>
                </div>
              )}
            </label>

            <label className="modal-field">
              <span>Link URL</span>
              <input type="url" value={form.link_url} onChange={handleChange('link_url')} placeholder="https://..." />
            </label>
          </div>

          <div className="banner-form-right">
            <div className="modal-row">
              <label className="modal-field">
                <span>Linked Offer</span>
                <select className="modal-select" value={form.linked_offer} onChange={handleChange('linked_offer')}>
                  <option value="">No linked offer</option>
                  {offers.map((o) => (
                    <option key={o.id} value={o.id}>{o.title}</option>
                  ))}
                </select>
              </label>
              <label className="modal-field">
                <span>Order</span>
                <input type="number" value={form.order} onChange={handleChange('order')} />
              </label>
            </div>

            <div className="modal-row">
              <label className="modal-field">
                <span>Start Date</span>
                <input type="datetime-local" value={form.start_date} onChange={handleChange('start_date')} />
              </label>
              <label className="modal-field">
                <span>End Date</span>
                <input type="datetime-local" value={form.end_date} onChange={handleChange('end_date')} />
              </label>
            </div>

            <label className="modal-field">
              <span>Branches</span>
              <BranchCheckDropdown selected={selectedBranchIds} onChange={setSelectedBranchIds} />
            </label>

            <label className="modal-check">
              <input type="checkbox" checked={form.is_active} onChange={handleChange('is_active')} />
              <span>Active</span>
            </label>
          </div>
        </div>

        <div className="modal-actions">
          <button type="button" className="small-button" onClick={onClose}>Cancel</button>
          <button type="submit" className="primary-button modal-submit" disabled={saving}>
            {saving ? 'Saving...' : mode === 'create' ? 'Create' : 'Update'}
          </button>
        </div>
      </form>
      {cropperSrc && (
        <ImageCropperModal
          imageSrc={cropperSrc}
          aspect={BANNER_IMAGE_ASPECT}
          onCrop={handleCropped}
          onCancel={() => setCropperSrc(null)}
        />
      )}
    </div>
  );
}

function BannersPage() {
  const [banners, setBanners] = useState([]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(1);
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const [loading, setLoading] = useState(true);
  const [modal, setModal] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);
  const [selectedIds, setSelectedIds] = useState([]);
  const [bulkBusy, setBulkBusy] = useState(false);
  const pageSize = 10;

  const fetchBanners = useCallback(async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ page, page_size: pageSize });
      if (search) params.set('search', search);
      const res = await apiFetch(`/api/banners/?${params}`);
      const data = await res.json();
      setBanners(data.results || []);
      setTotal(data.count || 0);
    } catch (err) {
      console.error('Failed to fetch banners', err);
    } finally {
      setLoading(false);
    }
  }, [page, search]);

  useEffect(() => { fetchBanners(); }, [fetchBanners]);

  const filteredBanners = banners.filter((b) => {
    if (statusFilter === 'active') return b.is_active === true;
    if (statusFilter === 'inactive') return b.is_active === false;
    return true;
  });

  const handleDelete = async () => {
    if (!deleteTarget) return;
    try {
      await apiFetch(`/api/banners/${deleteTarget.id}/`, {
        method: 'DELETE',
      });
      setDeleteTarget(null);
      fetchBanners();
    } catch (err) {
      console.error('Failed to delete banner', err);
    }
  };

  const toggleSelect = (id) => {
    setSelectedIds((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  };

  const handleToggleActive = async (banner) => {
    setBulkBusy(true);
    try {
      await apiFetch(`/api/banners/${banner.id}/`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ is_active: !banner.is_active }),
      });
      fetchBanners();
    } catch (err) {
      console.error('Failed to update banner', err);
    } finally {
      setBulkBusy(false);
    }
  };

  const selectAll = () => {
    const allIds = filteredBanners.map((b) => b.id);
    setSelectedIds((prev) => prev.length === allIds.length ? [] : allIds);
  };

  const handleBulkDelete = async () => {
    if (!selectedIds.length) return;
    setBulkBusy(true);
    await Promise.allSettled(
      selectedIds.map((id) =>
        apiFetch(`/api/banners/${id}/`, {
          method: 'DELETE',
        })
      )
    );
    setSelectedIds([]);
    setBulkBusy(false);
    fetchBanners();
  };

  const handleBulkSetActive = async () => {
    if (!selectedIds.length) return;
    setBulkBusy(true);
    await Promise.allSettled(
      selectedIds.map((id) =>
        apiFetch(`/api/banners/${id}/`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ is_active: true }),
        })
      )
    );
    setSelectedIds([]);
    setBulkBusy(false);
    setStatusFilter('all');
    setPage(1);
    fetchBanners();
  };

  const handleBulkSetInactive = async () => {
    if (!selectedIds.length) return;
    setBulkBusy(true);
    await Promise.allSettled(
      selectedIds.map((id) =>
        apiFetch(`/api/banners/${id}/`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ is_active: false }),
        })
      )
    );
    setSelectedIds([]);
    setBulkBusy(false);
    setStatusFilter('all');
    setPage(1);
    fetchBanners();
  };

  const totalPages = Math.ceil(total / pageSize);

  const filterTabs = [
    { key: 'all', label: 'All' },
    { key: 'active', label: 'Active' },
    { key: 'inactive', label: 'Inactive' },
  ];

  return (
    <>
      <div className="hero-row">
        <div>
          <p>Manage Banners</p>
          <h2>Banners</h2>
        </div>
        <div className="filters">
          <div className="search-field">
            <Search size={17} />
            <input type="text" placeholder="Search banners..." value={search} onChange={(e) => { setSearch(e.target.value); setPage(1); }} />
          </div>
          <button className="primary-button" onClick={() => setModal({ mode: 'create' })}>
            <Plus size={18} /> Add Banner
          </button>
        </div>
      </div>

      <section className="panel banners-panel">
        <div className="banners-toolbar">
          <div className="filter-tabs">
            {filterTabs.map((t) => (
              <button key={t.key} className={`filter-tab ${statusFilter === t.key ? 'active' : ''}`} onClick={() => { setStatusFilter(t.key); setPage(1); }}>
                {t.label}
              </button>
            ))}
          </div>
          {filteredBanners.length > 0 && (
            <label className="banners-select-all">
              <input
                type="checkbox"
                checked={selectedIds.length === filteredBanners.length}
                onChange={selectAll}
              />
              <span>Select all</span>
            </label>
          )}
        </div>

        {loading ? (
          <p className="table-loading">Loading banners...</p>
        ) : filteredBanners.length === 0 ? (
          search || statusFilter !== 'all' ? (
            <EmptyState icon={Search} title="No matching banners" description="No banners match your current filters. Try adjusting your search." />
          ) : (
            <EmptyState icon={Image} title="No banners yet" description="Create your first banner to promote offers across your branches." actionLabel="Add Banner" onAction={() => setModal({ mode: 'create' })} />
          )
        ) : (
          <div className="banners-grid">
            {filteredBanners.map((b) => (
              <article className={`banner-card ${selectedIds.includes(b.id) ? 'selected' : ''}`} key={b.id}>
                <div className="banner-card-media">
                  {b.image ? (
                    <img src={mediaUrl(b.image)} alt={b.title} className="banner-card-img" loading="lazy" />
                  ) : (
                    <div className="banner-card-img banner-card-img-empty">
                      <Image size={30} strokeWidth={1.4} />
                    </div>
                  )}
                  <label className="banner-card-check bulk-checkbox-cell">
                    <input
                      type="checkbox"
                      checked={selectedIds.includes(b.id)}
                      onChange={() => toggleSelect(b.id)}
                      aria-label={`Select ${b.title}`}
                    />
                  </label>
                  <span className={`banner-card-badge ${b.is_active ? '' : 'off'}`}>
                    {b.is_active ? 'Active' : 'Inactive'}
                  </span>
                </div>

                <div className="banner-card-body">
                  <strong className="banner-card-title" title={b.title}>{b.title}</strong>
                  <span className="banner-card-link" title={b.link_url || 'No link'}>
                    {b.link_url || 'No link'}
                  </span>
                </div>

                <div className="banner-card-foot">
                  <span className="banner-card-order">Order #{b.order ?? '—'}</span>
                  <span className="banner-card-actions">
                    {b.link_url && (
                      <a
                        href={b.link_url}
                        target="_blank"
                        rel="noreferrer"
                        aria-label="Open link"
                        title="Open link"
                      >
                        <ExternalLink size={15} />
                      </a>
                    )}
                    <button
                      onClick={() => setModal({ mode: 'edit', banner: b })}
                      aria-label="Edit"
                      title="Edit"
                    >
                      <Edit3 size={15} />
                    </button>
                    <button
                      onClick={() => handleToggleActive(b)}
                      disabled={bulkBusy}
                      aria-label={b.is_active ? 'Set Inactive' : 'Set Active'}
                      title={b.is_active ? 'Set Inactive' : 'Set Active'}
                    >
                      {b.is_active ? <EyeOff size={15} /> : <Eye size={15} />}
                    </button>
                    <button
                      onClick={() => setDeleteTarget(b)}
                      className="danger"
                      aria-label="Delete"
                      title="Delete"
                    >
                      <Trash2 size={15} />
                    </button>
                  </span>
                </div>
              </article>
            ))}
          </div>
        )}

        {totalPages > 1 && (
          <div className="pagination">
            <button disabled={page <= 1} onClick={() => setPage((p) => p - 1)}><ChevronLeft size={16} /> Prev</button>
            <span>Page {page} of {totalPages}</span>
            <button disabled={page >= totalPages} onClick={() => setPage((p) => p + 1)}>Next <ChevronRight size={16} /></button>
          </div>
        )}
      </section>

      {modal && (
        <BannerFormModal
          mode={modal.mode}
          banner={modal.banner}
          onClose={() => setModal(null)}
          onSaved={() => { setModal(null); fetchBanners(); }}
        />
      )}

      {deleteTarget && (
        <div className="modal-overlay" onClick={() => setDeleteTarget(null)}>
          <div className="modal-card confirm-card" onClick={(e) => e.stopPropagation()}>
            <h3>Delete Banner</h3>
            <p>Are you sure you want to delete <strong>{deleteTarget.title}</strong>?</p>
            <div className="modal-actions">
              <button className="small-button" onClick={() => setDeleteTarget(null)}>Cancel</button>
              <button className="danger-button" onClick={handleDelete}>Delete</button>
            </div>
          </div>
        </div>
      )}

      {selectedIds.length > 0 && (
        <BulkActionBar
          count={selectedIds.length}
          onDelete={handleBulkDelete}
          onSetActive={handleBulkSetActive}
          onSetInactive={handleBulkSetInactive}
          onClear={() => setSelectedIds([])}
          busy={bulkBusy}
        />
      )}
    </>
  );
}

function NotificationFormModal({ onClose, onSaved }) {
  const [form, setForm] = useState({
    title: '',
    body: '',
    type: 'COMMON',
    linked_offer: '',
    is_active: true,
    image: null,
  });
  const [sendToAll, setSendToAll] = useState(true);
  const [selectedBranchIds, setSelectedBranchIds] = useState([]);
  const [branches, setBranches] = useState([]);
  const [offers, setOffers] = useState([]);
  const [imagePreview, setImagePreview] = useState(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');

  useEffect(() => {
    Promise.all([
      apiFetch('/api/branches/').then((res) => res.json()),
      apiFetch('/api/offers/admin_all/').then((res) => res.json()),
    ]).then(([branchData, offerData]) => {
      setBranches(branchData.results || branchData || []);
      setOffers(offerData.results || offerData || []);
    }).catch(() => setError('Could not load branches and offers.'));
  }, []);

  const handleChange = (field) => (event) => {
    if (field === 'image') {
      const file = event.target.files?.[0] || null;
      setForm((prev) => ({ ...prev, image: file }));
      setImagePreview(file ? URL.createObjectURL(file) : null);
      return;
    }
    if (field === 'is_active') {
      setForm((prev) => ({ ...prev, is_active: event.target.checked }));
      return;
    }
    setForm((prev) => ({ ...prev, [field]: event.target.value }));
  };

  const handleSubmit = async (event) => {
    event.preventDefault();
    setSaving(true);
    setError('');

    try {
      const body = new FormData();
      body.append('title', form.title.trim());
      body.append('body', form.body.trim());
      body.append('type', form.type);
      body.append('is_active', form.is_active ? 'true' : 'false');
      if (form.linked_offer) body.append('linked_offer', form.linked_offer);
      if (form.image) body.append('image', form.image);
      if (!sendToAll) selectedBranchIds.forEach((id) => body.append('target_branches', id));

      const res = await apiFetch('/api/notifications/', { method: 'POST', body });
      if (!res.ok) {
        let message = 'Could not send notification.';
        try { message = JSON.stringify(await res.json()); } catch {}
        throw new Error(message);
      }
      onSaved();
    } catch (err) {
      setError(err.message || 'Could not send notification.');
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <form className="modal-card notification-modal-card" onClick={(event) => event.stopPropagation()} onSubmit={handleSubmit}>
        <div className="modal-head">
          <h3>Send Notification</h3>
          <button type="button" className="profile-close" onClick={onClose} aria-label="Close"><X size={18} /></button>
        </div>

        {error && <p className="notification-form-error">{error}</p>}

        <div className="notification-form-grid">
          <div>
            <label className="modal-field">
              <span>Title</span>
              <input value={form.title} onChange={handleChange('title')} maxLength="200" required />
            </label>
            <label className="modal-field">
              <span>Message</span>
              <textarea value={form.body} onChange={handleChange('body')} rows="5" required />
            </label>
            <label className="modal-field">
              <span>Image</span>
              <input type="file" accept="image/*" onChange={handleChange('image')} />
            </label>
            {imagePreview && <img src={imagePreview} alt="Notification preview" className="notification-preview" />}
          </div>

          <div>
            <label className="modal-field">
              <span>Notification Type</span>
              <select className="modal-select" value={form.type} onChange={handleChange('type')}>
                <option value="COMMON">Common</option>
                <option value="OFFER">Offer</option>
                <option value="PROMO">Promo</option>
              </select>
            </label>
            <label className="modal-field">
              <span>Linked Offer</span>
              <select className="modal-select" value={form.linked_offer} onChange={handleChange('linked_offer')}>
                <option value="">No linked offer</option>
                {offers.map((offer) => <option key={offer.id} value={offer.id}>{offer.title}</option>)}
              </select>
            </label>
            <label className="modal-check notification-audience-toggle">
              <input type="checkbox" checked={sendToAll} onChange={(event) => setSendToAll(event.target.checked)} />
              <span>Send to all users with FCM tokens</span>
            </label>
            {!sendToAll && (
              <label className="modal-field">
                <span>Target Branches</span>
                <BranchCheckDropdown selected={selectedBranchIds} onChange={setSelectedBranchIds} branches={branches} />
              </label>
            )}
            <label className="modal-check">
              <input type="checkbox" checked={form.is_active} onChange={handleChange('is_active')} />
              <span>Active</span>
            </label>
          </div>
        </div>

        <div className="modal-actions">
          <button type="button" className="small-button" onClick={onClose}>Cancel</button>
          <button type="submit" className="primary-button modal-submit" disabled={saving}>
            <Bell size={16} /> {saving ? 'Sending...' : 'Send Notification'}
          </button>
        </div>
      </form>
    </div>
  );
}

function NotificationsPage() {
  const [notifications, setNotifications] = useState([]);
  const [loading, setLoading] = useState(true);
  const [modalOpen, setModalOpen] = useState(false);
  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState('all');
  const [deleteTarget, setDeleteTarget] = useState(null);
  const [actionError, setActionError] = useState('');

  const fetchNotifications = useCallback(async () => {
    setLoading(true);
    try {
      const res = await apiFetch('/api/notifications/');
      if (!res.ok) throw new Error('Could not load notifications.');
      const data = await res.json();
      setNotifications(data.results || data || []);
    } catch (err) {
      setActionError(err.message || 'Could not load notifications.');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { fetchNotifications(); }, [fetchNotifications]);

  const handleDelete = async () => {
    if (!deleteTarget) return;
    try {
      const res = await apiFetch(`/api/notifications/${deleteTarget.id}/`, { method: 'DELETE' });
      if (!res.ok) throw new Error('Could not delete notification.');
      setDeleteTarget(null);
      fetchNotifications();
    } catch (err) {
      setActionError(err.message || 'Could not delete notification.');
    }
  };

  const visibleNotifications = notifications.filter((notification) => {
    const matchesType = typeFilter === 'all' || notification.type === typeFilter;
    const query = search.trim().toLowerCase();
    const matchesSearch = !query || notification.title.toLowerCase().includes(query) || notification.body.toLowerCase().includes(query);
    return matchesType && matchesSearch;
  });

  return (
    <>
      <div className="hero-row">
        <div>
          <p>Customer Messaging</p>
          <h2>Notifications</h2>
        </div>
        <div className="filters">
          <div className="search-field">
            <Search size={17} />
            <input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search notifications..." />
          </div>
          <button className="primary-button" onClick={() => setModalOpen(true)}>
            <Bell size={18} /> Send Notification
          </button>
        </div>
      </div>

      {actionError && <p className="notification-page-error">{actionError}</p>}

      <section className="panel notifications-panel">
        <div className="filter-tabs">
          {['all', 'COMMON', 'OFFER', 'PROMO'].map((type) => (
            <button key={type} className={`filter-tab ${typeFilter === type ? 'active' : ''}`} onClick={() => setTypeFilter(type)}>
              {type === 'all' ? 'All' : type[0] + type.slice(1).toLowerCase()}
            </button>
          ))}
        </div>

        {loading ? (
          <p className="table-loading">Loading notifications...</p>
        ) : visibleNotifications.length === 0 ? (
          <EmptyState icon={Bell} title="No notifications found" description="Send a notification to reach users with registered devices." actionLabel="Send Notification" onAction={() => setModalOpen(true)} />
        ) : (
          <div className="notifications-list">
            {visibleNotifications.map((notification) => (
              <article className="notification-row" key={notification.id}>
                <div className="notification-image">
                  {notification.image ? <img src={mediaUrl(notification.image)} alt="" /> : <Bell size={20} />}
                </div>
                <div className="notification-content">
                  <div className="notification-title-row">
                    <strong>{notification.title}</strong>
                    <span className={`notification-type type-${notification.type.toLowerCase()}`}>{notification.type}</span>
                  </div>
                  <p>{notification.body}</p>
                  <span>{new Date(notification.sent_at).toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })}</span>
                </div>
                <button className="notification-delete" onClick={() => setDeleteTarget(notification)} aria-label={`Delete ${notification.title}`} title="Delete">
                  <Trash2 size={16} />
                </button>
              </article>
            ))}
          </div>
        )}
      </section>

      {modalOpen && (
        <NotificationFormModal
          onClose={() => setModalOpen(false)}
          onSaved={() => { setModalOpen(false); fetchNotifications(); window.__showToast?.('Notification sent successfully.'); }}
        />
      )}

      {deleteTarget && (
        <div className="modal-overlay" onClick={() => setDeleteTarget(null)}>
          <div className="modal-card confirm-card" onClick={(event) => event.stopPropagation()}>
            <h3>Delete Notification</h3>
            <p>Delete <strong>{deleteTarget.title}</strong>? This also removes its user read history.</p>
            <div className="modal-actions">
              <button className="small-button" onClick={() => setDeleteTarget(null)}>Cancel</button>
              <button className="danger-button" onClick={handleDelete}>Delete</button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

function ActivityLogPage() {
  const [logs, setLogs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [entityFilter, setEntityFilter] = useState('all');
  const [actionFilter, setActionFilter] = useState('all');

  const fetchLogs = useCallback(async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams();
      if (entityFilter !== 'all') params.set('entity_type', entityFilter);
      if (actionFilter !== 'all') params.set('action', actionFilter);
      const qs = params.toString();
      const res = await apiFetch(`/api/activity-logs/${qs ? '?' + qs : ''}`);
      const data = await res.json();
      setLogs(Array.isArray(data) ? data : data.results || []);
    } catch (err) {
      console.error('Failed to fetch activity logs', err);
    } finally {
      setLoading(false);
    }
  }, [entityFilter, actionFilter]);

  useEffect(() => { fetchLogs(); }, [fetchLogs]);

  const formatTime = (iso) => {
    const d = new Date(iso);
    const now = new Date();
    const diffMs = now - d;
    const diffMin = Math.floor(diffMs / 60000);
    if (diffMin < 1) return 'Just now';
    if (diffMin < 60) return `${diffMin}m ago`;
    const diffHr = Math.floor(diffMin / 60);
    if (diffHr < 24) return `${diffHr}h ago`;
    const diffDay = Math.floor(diffHr / 24);
    if (diffDay < 7) return `${diffDay}d ago`;
    return d.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
  };

  const getActionColor = (action) => {
    if (action === 'CREATE') return 'action-create';
    if (action === 'UPDATE') return 'action-update';
    if (action === 'DELETE' || action === 'BULK_DELETE') return 'action-delete';
    if (action === 'LOGIN') return 'action-login';
    return 'action-bulk';
  };

  const entityTabs = [
    { key: 'all', label: 'All' },
    { key: 'branch', label: 'Branches' },
    { key: 'offer', label: 'Offers' },
    { key: 'banner', label: 'Banners' },
    { key: 'staff', label: 'Staff' },
    { key: 'account', label: 'Logins' },
  ];

  const actionTabs = [
    { key: 'all', label: 'All Actions' },
    { key: 'CREATE', label: 'Created' },
    { key: 'UPDATE', label: 'Updated' },
    { key: 'DELETE', label: 'Deleted' },
    { key: 'BULK_DELETE', label: 'Bulk Deleted' },
    { key: 'BULK_UPDATE', label: 'Bulk Updated' },
    { key: 'LOGIN', label: 'Login' },
  ];

  return (
    <>
      <div className="hero-row">
        <div>
          <p>Track Actions</p>
          <h2>Activity Log</h2>
        </div>
      </div>

      <section className="panel activity-panel">
        <div className="activity-filters-row">
          <div className="filter-tabs">
            {entityTabs.map((t) => (
              <button key={t.key} className={`filter-tab ${entityFilter === t.key ? 'active' : ''}`} onClick={() => setEntityFilter(t.key)}>
                {t.label}
              </button>
            ))}
          </div>
          <div className="activity-action-filter">
            <label>Action:</label>
            <select className="modal-select" value={actionFilter} onChange={(e) => setActionFilter(e.target.value)}>
              {actionTabs.map((t) => (
                <option key={t.key} value={t.key}>{t.label}</option>
              ))}
            </select>
          </div>
        </div>

        {loading ? (
          <p className="table-loading">Loading activity logs...</p>
        ) : logs.length === 0 ? (
          <EmptyState icon={Clock3} title="No activity yet" description="Actions performed by admins will appear here." />
        ) : (
          <div className="activity-table">
            <div className="activity-table-head">
              <span>Time</span>
              <span>User</span>
              <span>Action</span>
              <span>Type</span>
              <span>Entity</span>
            </div>
            {logs.map((log) => (
              <div className="activity-table-row" key={log.id}>
                <span className="activity-time">{formatTime(log.created_at)}</span>
                <div className="activity-user">
                  <span className="activity-user-avatar">{log.user_name?.[0] || 'U'}</span>
                  <span>{log.user_name}</span>
                </div>
                <span className={`activity-action-badge ${getActionColor(log.action)}`}>{log.action_display}</span>
                <span className="activity-entity-type">{log.entity_type}</span>
                <span className="activity-entity-name">{log.entity_name || `#${log.entity_id}`}</span>
              </div>
            ))}
          </div>
        )}
      </section>
    </>
  );
}


function App() {
  const [authData, setAuthData] = useState(() => getSession());
  const [activePage, setActivePage] = useState('Dashboard');
  const [showSearch, setShowSearch] = useState(false);
  const [selectedBranch, setSelectedBranch] = useState(null);
  const [navItems, setNavItems] = useState(loadNavItems);
  const [editMode, setEditMode] = useState(false);
  const [toast, setToast] = useState(null);

  useEffect(() => subscribeToSession(setAuthData), []);

  useEffect(() => {
    if (getSession()) refreshSession();
  }, []);

  useEffect(() => {
    window.__showToast = (msg) => setToast(msg);
    return () => { delete window.__showToast; };
  }, []);

  const navigateTo = useCallback((page, options = {}) => {
    setSelectedBranch(page === 'Branches' ? options.branch || null : null);
    setActivePage(page);
    window.scrollTo({ top: 0 });
    const content = document.querySelector('.content');
    if (content) content.scrollTop = 0;
  }, []);

  const handleReorder = useCallback((newItems) => {
    if (newItems === null) {
      setEditMode(false);
      return;
    }
    setNavItems(newItems);
    saveNavItems(newItems);
  }, []);

  const toggleEditMode = useCallback(() => {
    setEditMode((prev) => !prev);
  }, []);

  const publicFormMatch = window.location.pathname.match(/^\/forms\/([0-9a-f-]+)\/?$/i);
  if (publicFormMatch) {
    return <PublicFormPage apiFetch={apiFetch} publicId={publicFormMatch[1]} apiBase={API_BASE_URL} />;
  }

  if (!authData) {
    return <LoginPage onLogin={(data) => setSession(data)} />;
  }

  const pages = {
    Dashboard: <DashboardPage authData={authData} onNavigate={navigateTo} />,
    Branches: <BranchesPage selectedBranch={selectedBranch} onClearSelectedBranch={() => setSelectedBranch(null)} />,
    Offers: <OffersPage />,
    Banners: <BannersPage />,
    Notifications: <NotificationsPage />,
    Customers: <AccountUsersPage />,
    Forms: <CustomFormsPage apiFetch={apiFetch} apiBase={API_BASE_URL} />,
    Users: <UsersPage />,
    Activity: <ActivityLogPage />,
  };

  return (
    <main className="page">
      <div className="app-shell">
        <Rail authData={authData} onOpenSearch={() => setShowSearch(true)} onLogout={() => setSession(null)} editMode={editMode} onToggleEdit={toggleEditMode} />
        <div className="workspace">
          <Sidebar activePage={activePage} onNavigate={navigateTo} navItems={navItems} editMode={editMode} onToggleEdit={toggleEditMode} onReorder={handleReorder} />
          <section className="content">
            <header className="topbar">
              <div />
              <div />
            </header>
            {pages[activePage] || <DashboardPage authData={authData} onNavigate={navigateTo} />}
      </section>
        </div>
      </div>
      {showSearch && (
        <GlobalSearch
          onClose={() => setShowSearch(false)}
          onNavigate={navigateTo}
        />
      )}
      {toast && <Toast message={toast} onClose={() => setToast(null)} />}
    </main>
  );
}

createRoot(document.getElementById('root')).render(<App />);
