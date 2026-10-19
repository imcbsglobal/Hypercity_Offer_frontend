import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { Check, ChevronLeft, ChevronRight, Clipboard, Download, Eye, FileText, GripVertical, Plus, Send, Star, Trash2, X } from 'lucide-react';
import './customForms.css';

const FIELD_TYPES = [
  ['SHORT_TEXT', 'Short text'], ['LONG_TEXT', 'Long text'], ['EMAIL', 'Email'], ['PHONE', 'Phone'],
  ['NUMBER', 'Number'], ['DATE', 'Date'], ['SINGLE_CHOICE', 'Single choice'],
  ['MULTIPLE_CHOICE', 'Multiple choice'], ['DROPDOWN', 'Dropdown'], ['RATING', 'Rating'],
];
const CHOICE_TYPES = new Set(['SINGLE_CHOICE', 'MULTIPLE_CHOICE', 'DROPDOWN']);
const ACCENT_PRESETS = ['#1A7A48', '#2563EB', '#7C3AED', '#DB2777', '#EA580C', '#0F766E'];
const BACKGROUND_PRESETS = ['#FFFFFF', '#F8FAFC', '#ECFDF5', '#EFF6FF', '#FFF7ED', '#FDF2F8'];

function choiceOptions(fieldType, choices) {
  const options = Array.isArray(choices) ? choices : [];
  const minimum = fieldType === 'SINGLE_CHOICE' ? 1 : 2;
  return options.length >= minimum ? options : [...options, ...['Option 1', 'Option 2'].slice(options.length, minimum)];
}

function blankField(type, position) {
  const label = FIELD_TYPES.find(([value]) => value === type)?.[1] || 'Question';
  return {
    field_key: crypto.randomUUID(), position, field_type: type, label, help_text: '', placeholder: '', required: false,
    choices: CHOICE_TYPES.has(type) ? choiceOptions(type, []) : [], validation: type === 'RATING' ? { min: 1, max: 5 } : {},
  };
}

function formPath(apiBase, path) {
  return `${apiBase || ''}${path}`;
}

function mediaUrl(apiBase, path) {
  if (!path || /^(?:https?:|blob:|data:)/i.test(path)) return path || '';
  return formPath(apiBase, path);
}

function publicFormLink(form) {
  return form.share_url?.startsWith('http') ? form.share_url : `${window.location.origin}${form.share_url}`;
}

function QuestionInput({ field, value, onChange, disabled = false }) {
  const id = `field-${field.field_key}`;
  const required = field.required ? <span className="public-required">*</span> : null;
  if (field.field_type === 'LONG_TEXT') return <textarea id={id} value={value || ''} placeholder={field.placeholder} onChange={(e) => onChange(e.target.value)} disabled={disabled} />;
  if (field.field_type === 'SINGLE_CHOICE') return <div className="public-options">{field.choices.map((choice) => <label key={choice}><input type="radio" name={id} checked={value === choice} onChange={() => onChange(choice)} disabled={disabled} /> {choice}</label>)}</div>;
  if (field.field_type === 'MULTIPLE_CHOICE') return <div className="public-options">{field.choices.map((choice) => <label key={choice}><input type="checkbox" checked={(value || []).includes(choice)} onChange={(e) => onChange(e.target.checked ? [...(value || []), choice] : (value || []).filter((item) => item !== choice))} disabled={disabled} /> {choice}</label>)}</div>;
  if (field.field_type === 'DROPDOWN') return <select id={id} value={value || ''} onChange={(e) => onChange(e.target.value)} disabled={disabled}><option value="">Select an option</option>{field.choices.map((choice) => <option key={choice}>{choice}</option>)}</select>;
  if (field.field_type === 'RATING') {
    const min = field.validation?.min || 1;
    const max = field.validation?.max || 5;
    return <div className="rating-input">{Array.from({ length: max - min + 1 }, (_, index) => min + index).map((rating) => <button type="button" key={rating} className={Number(value) >= rating ? 'selected' : ''} onClick={() => onChange(rating)} disabled={disabled} aria-label={`${rating} stars`}><Star size={22} fill={Number(value) >= rating ? 'currentColor' : 'none'} /></button>)}</div>;
  }
  const typeMap = { EMAIL: 'email', PHONE: 'tel', NUMBER: 'number', DATE: 'date' };
  return <input id={id} type={typeMap[field.field_type] || 'text'} value={value || ''} placeholder={field.placeholder} onChange={(e) => onChange(e.target.value)} disabled={disabled} />;
}

function OptionsEditor({ fieldType, options, onChange, disabled }) {
  const editableOptions = choiceOptions(fieldType, options);
  const minimumOptions = fieldType === 'SINGLE_CHOICE' ? 1 : 2;
  const showRemoveButtons = fieldType !== 'SINGLE_CHOICE' || editableOptions.length > 1;
  const updateOption = (index, value) => onChange(editableOptions.map((option, optionIndex) => optionIndex === index ? value : option));
  const removeOption = (index) => onChange(editableOptions.filter((_, optionIndex) => optionIndex !== index));
  return <div className="options-editor"><span>Options</span>{editableOptions.map((option, index) => <div className={showRemoveButtons ? 'option-editor-row' : 'option-editor-row option-editor-row-fixed'} key={index}><input value={option} disabled={disabled} onChange={(event) => updateOption(index, event.target.value)} aria-label={`Option ${index + 1}`} />{showRemoveButtons && <button type="button" onClick={() => removeOption(index)} disabled={disabled || editableOptions.length <= minimumOptions} title="Remove option"><X size={14} /></button>}</div>)}{!disabled && fieldType !== 'SINGLE_CHOICE' && <button type="button" className="add-option-button" onClick={() => onChange([...editableOptions, `Option ${editableOptions.length + 1}`])}><Plus size={14} /> Add option</button>}</div>;
}

function BuilderPreview({ form }) {
  return <aside className="form-preview" style={{ '--form-accent': form.accent_color || '#1A7A48', backgroundColor: form.background_color || '#FFFFFF' }}>
    <span className="preview-label"><Eye size={14} /> Live preview</span>
    {form.logoPreview && <img src={form.logoPreview} alt="Form logo" className="preview-logo" />}
    <h3>{form.title || 'Untitled form'}</h3>
    {form.description && <p>{form.description}</p>}
    <div className="preview-fields">{form.fields.length ? form.fields.map((field) => <div className="public-question" key={field.field_key}><label>{field.label || 'Untitled question'} {field.required && <span className="public-required">*</span>}</label>{field.help_text && <small>{field.help_text}</small>}<QuestionInput field={field} value={field.field_type === 'MULTIPLE_CHOICE' ? [] : ''} onChange={() => {}} disabled /></div>) : <p className="preview-empty">Add questions from the left panel.</p>}</div>
  </aside>;
}

function FormBuilder({ apiFetch, apiBase, initialForm, onBack, onSaved }) {
  const [form, setForm] = useState(() => ({ ...initialForm, fields: initialForm.fields || [], logo: null, logoPreview: mediaUrl(apiBase, initialForm.logo_url) }));
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');
  const [dragIndex, setDragIndex] = useState(null);
  const [showSetup, setShowSetup] = useState(!initialForm.id);
  const isDraft = !form.id || form.status === 'DRAFT';
  const isNew = !form.id;
  const validAccent = /^#[0-9A-Fa-f]{6}$/.test(form.accent_color || '') ? form.accent_color : '#1A7A48';
  const validBackground = /^#[0-9A-Fa-f]{6}$/.test(form.background_color || '') ? form.background_color : '#FFFFFF';

  const updateForm = (key, value) => setForm((current) => ({ ...current, [key]: value }));
  const updateField = (index, patch) => setForm((current) => ({ ...current, fields: current.fields.map((field, currentIndex) => currentIndex === index ? { ...field, ...patch } : field) }));
  const addField = (type) => setForm((current) => ({ ...current, fields: [...current.fields, blankField(type, current.fields.length)] }));
  const removeField = (index) => setForm((current) => ({ ...current, fields: current.fields.filter((_, currentIndex) => currentIndex !== index).map((field, position) => ({ ...field, position })) }));
  const reorder = (targetIndex) => {
    if (dragIndex === null || dragIndex === targetIndex) return;
    setForm((current) => {
      const fields = [...current.fields];
      const [moved] = fields.splice(dragIndex, 1);
      fields.splice(targetIndex, 0, moved);
      return { ...current, fields: fields.map((field, position) => ({ ...field, position })) };
    });
    setDragIndex(null);
  };

  const save = async () => {
    setSaving(true); setError('');
    try {
      const body = new FormData();
      body.append('title', form.title.trim()); body.append('description', form.description || '');
      body.append('accent_color', form.accent_color || '#1A7A48'); body.append('background_color', form.background_color || '#FFFFFF'); body.append('confirmation_message', form.confirmation_message || 'Thanks for your response!');
      body.append('fields', JSON.stringify(form.fields.map(({ logo, logoPreview, ...field }) => field)));
      if (form.logo instanceof File) body.append('logo', form.logo);
      const path = form.id ? `/api/custom-forms/${form.id}/` : '/api/custom-forms/';
      const response = await apiFetch(path, { method: form.id ? 'PATCH' : 'POST', body });
      const data = await response.json();
      if (!response.ok) throw new Error(JSON.stringify(data));
      setForm((current) => ({ ...data, logo: null, logoPreview: mediaUrl(apiBase, data.logo_url) || current.logoPreview }));
      onSaved(data);
      window.__showToast?.('Draft saved.');
      return data;
    } catch (err) { setError(err.message || 'Could not save form.'); return null; } finally { setSaving(false); }
  };

  const changeStatus = async (action) => {
    const saved = action === 'publish' ? await save() : form;
    if (!saved) return;
    setSaving(true); setError('');
    try {
      const response = await apiFetch(`/api/custom-forms/${saved.id}/${action}/`, { method: 'POST' });
      const data = await response.json();
      if (!response.ok) throw new Error(JSON.stringify(data));
      setForm((current) => ({ ...current, ...data, logoPreview: mediaUrl(apiBase, data.logo_url) || current.logoPreview }));
      onSaved(data); window.__showToast?.(action === 'publish' ? 'Form published.' : `Form ${action}ed.`);
    } catch (err) { setError(err.message || 'Could not change form status.'); } finally { setSaving(false); }
  };

  const copyLink = async () => {
    await navigator.clipboard.writeText(publicFormLink(form)); window.__showToast?.('Public link copied.');
  };
  const viewPublicForm = () => window.open(publicFormLink(form), '_blank', 'noopener,noreferrer');

  return <div className="forms-builder">
    <div className="forms-builder-head"><button className="icon-button" onClick={onBack} title="Back to forms"><ChevronLeft size={19} /></button><div><p>Custom Form Builder</p><h2>{form.title || 'New form'}</h2></div><div className="builder-actions">{form.status === 'PUBLISHED' && <button className="small-button" onClick={() => changeStatus('close')} disabled={saving}>Close form</button>}{form.status === 'CLOSED' && <button className="small-button" onClick={() => changeStatus('reopen')} disabled={saving}>Reopen form</button>}{form.status === 'PUBLISHED' && <><button className="small-button" onClick={viewPublicForm}><Eye size={15} /> View public form</button><button className="small-button" onClick={copyLink}><Clipboard size={15} /> Copy link</button></>}{!isNew && isDraft && !showSetup && <><button className="small-button" onClick={save} disabled={saving}>Save draft</button><button className="primary-button" onClick={() => changeStatus('publish')} disabled={saving}><Send size={16} /> Publish</button></>}</div></div>
    {error && <p className="forms-error">{error}</p>}
    {form.status === 'PUBLISHED' && <div className="share-link"><span>Public link</span><code>{form.share_url}</code><button onClick={copyLink} title="Copy public link"><Clipboard size={15} /></button></div>}
    <div className="builder-layout">
      {showSetup ? <section className="builder-editor"><div className="form-settings"><h3>Step 1: Form settings</h3><label>Title<input value={form.title || ''} onChange={(e) => updateForm('title', e.target.value)} disabled={!isDraft} required /></label><label>Description<textarea value={form.description || ''} onChange={(e) => updateForm('description', e.target.value)} disabled={!isDraft} /></label><div className="form-settings-row"><label>Accent color<div className="accent-color-control"><input className="accent-swatch-input" type="color" value={validAccent} onChange={(e) => updateForm('accent_color', e.target.value.toUpperCase())} disabled={!isDraft} aria-label="Choose accent color" /><input className="accent-hex-input" value={form.accent_color || ''} maxLength="7" onChange={(e) => updateForm('accent_color', e.target.value)} disabled={!isDraft} aria-label="Accent color hex value" /><div className="accent-presets">{ACCENT_PRESETS.map((color) => <button type="button" key={color} title={color} aria-label={`Use ${color}`} className={validAccent.toUpperCase() === color ? 'active' : ''} style={{ backgroundColor: color }} onClick={() => updateForm('accent_color', color)} disabled={!isDraft} />)}</div></div></label><label>Form background<div className="accent-color-control"><input className="accent-swatch-input" type="color" value={validBackground} onChange={(e) => updateForm('background_color', e.target.value.toUpperCase())} disabled={!isDraft} aria-label="Choose form background color" /><input className="accent-hex-input" value={form.background_color || ''} maxLength="7" onChange={(e) => updateForm('background_color', e.target.value)} disabled={!isDraft} aria-label="Form background hex value" /><div className="accent-presets">{BACKGROUND_PRESETS.map((color) => <button type="button" key={color} title={color} aria-label={`Use ${color}`} className={validBackground.toUpperCase() === color ? 'active' : ''} style={{ backgroundColor: color }} onClick={() => updateForm('background_color', color)} disabled={!isDraft} />)}</div></div></label></div><label>Logo<input type="file" accept="image/*" disabled={!isDraft} onChange={(e) => { const file = e.target.files?.[0]; if (file) setForm((current) => ({ ...current, logo: file, logoPreview: URL.createObjectURL(file) })); }} /></label><label>Confirmation message<textarea value={form.confirmation_message || ''} onChange={(e) => updateForm('confirmation_message', e.target.value)} disabled={!isDraft} /></label>{isDraft && <div className="setup-actions"><button className="primary-button" onClick={async () => { const saved = await save(); if (saved) setShowSetup(false); }} disabled={saving}>{saving ? 'Saving...' : <>Next: Add questions <ChevronRight size={16} /></>}</button></div>}</div></section> : <section className="builder-editor">{isDraft && <div className="field-palette"><h3>Step 2: Add questions</h3>{FIELD_TYPES.map(([type, label]) => <button key={type} onClick={() => addField(type)}><Plus size={14} /> {label}</button>)}</div>}<div className="builder-questions"><h3>Questions</h3>{form.fields.map((field, index) => <article className="builder-question" key={field.field_key} draggable={isDraft} onDragStart={() => setDragIndex(index)} onDragOver={(e) => e.preventDefault()} onDrop={() => reorder(index)}><GripVertical size={18} className="question-grip" /><div className="question-settings"><div className="question-row"><label>Type<select value={field.field_type} disabled={!isDraft} onChange={(e) => updateField(index, { field_type: e.target.value, choices: CHOICE_TYPES.has(e.target.value) ? choiceOptions(e.target.value, field.choices) : [] })}>{FIELD_TYPES.map(([type, label]) => <option key={type} value={type}>{label}</option>)}</select></label><label className="required-toggle"><input type="checkbox" checked={field.required} disabled={!isDraft} onChange={(e) => updateField(index, { required: e.target.checked })} /> Required</label></div><label>Question<input value={field.label} disabled={!isDraft} onChange={(e) => updateField(index, { label: e.target.value })} /></label><label>Help text<input value={field.help_text || ''} disabled={!isDraft} onChange={(e) => updateField(index, { help_text: e.target.value })} /></label>{!CHOICE_TYPES.has(field.field_type) && field.field_type !== 'RATING' && <label>Placeholder<input value={field.placeholder || ''} disabled={!isDraft} onChange={(e) => updateField(index, { placeholder: e.target.value })} /></label>}{CHOICE_TYPES.has(field.field_type) && <OptionsEditor fieldType={field.field_type} options={field.choices || []} disabled={!isDraft} onChange={(choices) => updateField(index, { choices })} />}{field.field_type === 'RATING' && <div className="question-row"><label>Min<input type="number" min="1" max="9" value={field.validation?.min || 1} disabled={!isDraft} onChange={(e) => updateField(index, { validation: { ...field.validation, min: Number(e.target.value) } })} /></label><label>Max<input type="number" min="2" max="10" value={field.validation?.max || 5} disabled={!isDraft} onChange={(e) => updateField(index, { validation: { ...field.validation, max: Number(e.target.value) } })} /></label></div>}</div>{isDraft && <button className="question-delete" onClick={() => removeField(index)} title="Remove question"><Trash2 size={16} /></button>}</article>)}{!form.fields.length && <p className="forms-empty">Choose a question type to start building.</p>}</div>{isDraft && <div className="setup-actions"><button className="primary-button" onClick={() => setShowSetup(true)}><ChevronLeft size={16} /> Back to form settings</button></div>}</section>}
      <BuilderPreview form={form} />
    </div>
  </div>;
}

function ResponsesPanel({ apiFetch, form, onClose }) {
  const [responses, setResponses] = useState([]); const [loading, setLoading] = useState(true); const [from, setFrom] = useState(''); const [to, setTo] = useState(''); const [selected, setSelected] = useState(null);
  const load = useCallback(async () => { setLoading(true); const params = new URLSearchParams(); if (from) params.set('submitted_after', from); if (to) params.set('submitted_before', to); const response = await apiFetch(`/api/custom-forms/${form.id}/responses/?${params}`); const data = await response.json(); setResponses(data.results || data || []); setLoading(false); }, [apiFetch, form.id, from, to]);
  useEffect(() => { load(); }, [load]);
  const exportCsv = async () => { const response = await apiFetch(`/api/custom-forms/${form.id}/export/`); const blob = await response.blob(); const url = URL.createObjectURL(blob); const anchor = document.createElement('a'); anchor.href = url; anchor.download = `${form.title}-responses.csv`; anchor.click(); URL.revokeObjectURL(url); };
  return <div className="responses-drawer"><div className="responses-head"><div><p>Submitted data</p><h2>{form.title}</h2></div><button className="icon-button" onClick={onClose}><X size={19} /></button></div><div className="responses-tools"><label>From<input type="date" value={from} onChange={(e) => setFrom(e.target.value)} /></label><label>To<input type="date" value={to} onChange={(e) => setTo(e.target.value)} /></label><button className="small-button" onClick={exportCsv}><Download size={15} /> Export CSV</button></div>{loading ? <p className="forms-empty">Loading responses...</p> : <div className="response-list">{responses.length ? responses.map((response, index) => <button key={response.id} onClick={() => setSelected(response)}><span>Response #{index + 1}</span><small>{new Date(response.submitted_at).toLocaleString()}</small></button>) : <p className="forms-empty">No responses yet.</p>}</div>}{selected && <div className="response-detail"><div><strong>Response #{responses.findIndex((item) => item.id === selected.id) + 1}</strong><button onClick={() => setSelected(null)}><X size={15} /></button></div>{selected.schema_snapshot.map((field) => <dl key={field.field_key}><dt>{field.label}</dt><dd>{Array.isArray(selected.answers[field.field_key]) ? selected.answers[field.field_key].join(', ') : String(selected.answers[field.field_key] ?? '-')}</dd></dl>)}</div>}</div>;
}

export function CustomFormsPage({ apiFetch, apiBase }) {
  const [forms, setForms] = useState([]); const [loading, setLoading] = useState(true); const [editing, setEditing] = useState(null); const [responsesForm, setResponsesForm] = useState(null);
  const load = useCallback(async () => { setLoading(true); try { const response = await apiFetch('/api/custom-forms/'); const data = await response.json(); setForms(data.results || data || []); } finally { setLoading(false); } }, [apiFetch]);
  useEffect(() => { load(); }, [load]);
  const remove = async (form) => { if (!window.confirm(`Delete "${form.title}"? This permanently removes all responses.`)) return; await apiFetch(`/api/custom-forms/${form.id}/`, { method: 'DELETE' }); load(); };
  const viewPublicForm = (form) => window.open(publicFormLink(form), '_blank', 'noopener,noreferrer');
  if (editing) return <FormBuilder apiFetch={apiFetch} apiBase={apiBase} initialForm={editing} onBack={() => { setEditing(null); load(); }} onSaved={(saved) => setEditing((current) => current ? { ...current, ...saved } : saved)} />;
  return <><div className="hero-row"><div><p>Public data collection</p><h2>Forms</h2></div><button className="primary-button" onClick={() => setEditing({ title: '', description: '', accent_color: '#1A7A48', background_color: '#FFFFFF', confirmation_message: 'Thanks for your response!', status: 'DRAFT', fields: [] })}><Plus size={18} /> Create form</button></div><section className="panel forms-panel">{loading ? <p className="forms-empty">Loading forms...</p> : forms.length ? <div className="forms-list">{forms.map((form) => <article key={form.id}><div className="form-list-icon" style={{ backgroundColor: form.accent_color }}><FileText size={21} /></div><div><strong>{form.title}</strong><span>{form.response_count || 0} responses</span></div><span className={`form-status status-${form.status.toLowerCase()}`}>{form.status}</span><div className="form-list-actions"><button onClick={() => setEditing(form)}>Open</button>{form.status === 'PUBLISHED' && <button onClick={() => viewPublicForm(form)} title="View public form"><Eye size={15} /> View</button>}<button onClick={() => setResponsesForm(form)}>Responses</button><button className="danger-text" onClick={() => remove(form)} title="Delete form"><Trash2 size={16} /></button></div></article>)}</div> : <div className="forms-empty"><FileText size={42} /><h3>No forms yet</h3><p>Create a public form and share its generated link.</p><button className="primary-button" onClick={() => setEditing({ title: '', description: '', accent_color: '#1A7A48', background_color: '#FFFFFF', confirmation_message: 'Thanks for your response!', status: 'DRAFT', fields: [] })}><Plus size={16} /> Create form</button></div>}</section>{responsesForm && <ResponsesPanel apiFetch={apiFetch} form={responsesForm} onClose={() => setResponsesForm(null)} />}</>;
}

export function PublicFormPage({ apiFetch, publicId, apiBase }) {
  const [form, setForm] = useState(null); const [answers, setAnswers] = useState({}); const [errors, setErrors] = useState({}); const [state, setState] = useState('loading'); const [message, setMessage] = useState('');
  useEffect(() => { apiFetch(`/api/public/forms/${publicId}/`).then(async (response) => { if (!response.ok) { setState('unavailable'); return; } setForm(await response.json()); setState('ready'); }).catch(() => setState('unavailable')); }, [apiFetch, publicId]);
  const submit = async (event) => { event.preventDefault(); setErrors({}); const response = await apiFetch(`/api/public/forms/${publicId}/responses/`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ answers }) }); const data = await response.json(); if (!response.ok) { setErrors(data.answers || data); return; } setMessage(data.message); setState('complete'); };
  if (state === 'loading') return <main className="public-form-page"><p>Loading form...</p></main>;
  if (state === 'unavailable') return <main className="public-form-page"><section className="public-form-card unavailable"><h1>This form is unavailable</h1><p>It may be closed, unpublished, or no longer exist.</p></section></main>;
  if (state === 'complete') return <main className="public-form-page"><section className="public-form-card complete"><Check size={38} /><h1>Response submitted</h1><p>{message}</p></section></main>;
  const logo = form.logo_url ? formPath(apiBase, form.logo_url) : '';
  return <main className="public-form-page" style={{ '--form-accent': form.accent_color }}><form className="public-form-card" style={{ backgroundColor: form.background_color || '#FFFFFF' }} onSubmit={submit}>{logo && <img src={logo} alt="Form logo" className="public-form-logo" />}<h1>{form.title}</h1>{form.description && <p className="public-description">{form.description}</p>}{form.fields.map((field) => <div className="public-question" key={field.field_key}><label htmlFor={`field-${field.field_key}`}>{field.label} {field.required && <span className="public-required">*</span>}</label>{field.help_text && <small>{field.help_text}</small>}<QuestionInput field={field} value={answers[field.field_key]} onChange={(value) => setAnswers((current) => ({ ...current, [field.field_key]: value }))} />{errors[field.field_key] && <em>{Array.isArray(errors[field.field_key]) ? errors[field.field_key].join(' ') : errors[field.field_key]}</em>}</div>)}<button className="public-submit" type="submit">Submit response</button></form></main>;
}
