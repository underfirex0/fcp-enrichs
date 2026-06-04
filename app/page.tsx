/* eslint-disable @typescript-eslint/no-explicit-any */
'use client'
import { useState, useRef, useEffect, useCallback } from 'react'
import {
  Settings, Table2, ChevronDown, ChevronRight, Play, Pause, Square,
  Download, Search, X, Check, Plus, Minus, RefreshCw, Database,
  User, Mail, Phone, Globe, TrendingUp, Users, BarChart3, Layers,
  ExternalLink, AlertCircle,
} from 'lucide-react'

// ─── Types ────────────────────────────────────────────────────────────────
interface Company {
  code: string; name: string; city: string
  ca?: string; caExp?: string; effectif?: string; activite?: string
  sectorsChecked?: string[]   // Filières cochées dans l'Excel (à vérifier)
}
interface EnrichedResult extends Company {
  dirigeant: string | null; email_dirigeant: string | null
  mobile_dirigeant: string | null; site_web: string | null
  ca_montant: string | null
  tranche_ca: string | null; ca_export: string | null
  part_export: string | null; effectif_tranche: string | null
  filieres_confirmees: string[]; filieres_rejetees: string[]
  filieres_ajoutees: string[]; filiere_principale: string | null
  secteur_reel: string | null
  type_entreprise: string | null; confiance: number
  raison: string; sources: string[]; processed_at: string
  error?: string
}

// ─── Constants ────────────────────────────────────────────────────────────
const FILIERE_COLORS: Record<string, { c: string; bg: string }> = {
  "Engrais & fertilisants":            { c: '#2A7A52', bg: '#E0F0E8' },
  "Produits de commodité":             { c: '#2850A0', bg: '#E0E8F5' },
  "Traitement de l'eau":               { c: '#1878A0', bg: '#D8EEF8' },
  "Cosmétiques & bien-être":           { c: '#A03080', bg: '#F5E0F0' },
  "Détergents & produits de nettoyage":{ c: '#B87020', bg: '#F5EDDC' },
  "Peintures, colles & encres":        { c: '#7040A0', bg: '#EDE8F5' },
  "Pigments & colorants":              { c: '#C84020', bg: '#F5E0DC' },
  "Polymères & résines":               { c: '#507070', bg: '#E0ECEC' },
  "Verres":                            { c: '#1878A0', bg: '#D8EAF5' },
  "Phytosanitaires":                   { c: '#408028', bg: '#E4F0DC' },
  "Gaz industriels & médicaux":        { c: '#6040A0', bg: '#EAE4F5' },
  "Hors secteur FCP":                  { c: '#808080', bg: '#F0F0F0' },
}
const FILIERES_LIST = Object.keys(FILIERE_COLORS)

const TYPE_COLORS: Record<string, { c: string; bg: string }> = {
  "Fabricant":               { c: '#2A7A52', bg: '#E0F0E8' },
  "Fabricant-Distributeur":  { c: '#1878A0', bg: '#D8EAF5' },
  "Distributeur":            { c: '#2850A0', bg: '#E0E8F5' },
  "Importateur":             { c: '#7040A0', bg: '#EDE8F5' },
  "Importateur-Distributeur":{ c: '#604888', bg: '#EAE4F5' },
  "Agent/Représentant":      { c: '#B87020', bg: '#F5EDDC' },
  "Négoce":                  { c: '#C84020', bg: '#F5E0DC' },
  "Prestataire de services": { c: '#607060', bg: '#E8EEE8' },
}
const TYPES_LIST = Object.keys(TYPE_COLORS)

// ─── Utils ────────────────────────────────────────────────────────────────
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))
function save(k: string, v: unknown) { try { localStorage.setItem(k, JSON.stringify(v)) } catch {} }
function load<T>(k: string, fb: T): T {
  try { const v = localStorage.getItem(k); return v !== null ? JSON.parse(v) : fb } catch { return fb }
}

function parseRaw(raw: string): Company[] {
  return raw.split('\n')
    .map(l => l.trim()).filter(Boolean)
    .map(l => {
      const parts = l.split('\t').map(p => p.trim())
      return {
        code:     parts[0] || '',
        name:     parts[1] || '',
        city:     parts[2] || '',
        ca:       parts[3] || undefined,
        caExp:    parts[4] || undefined,
        effectif: parts[5] || undefined,
        activite: parts[6] || undefined,
      }
    })
    .filter(c => c.name && c.code)
}

// ─── Excel sector column mapping ──────────────────────────────────────────
const SECTOR_MAP: { match: string[]; filiere: string }[] = [
  { match: ['engrais', 'fertilisant', 'angrais'],            filiere: 'Engrais & fertilisants' },
  { match: ['commodité', 'commodite'],                        filiere: 'Produits de commodité' },
  { match: ['traitement', "l'eau"],                           filiere: "Traitement de l'eau" },
  { match: ['cosmétique', 'cosmetique', 'bien-être'],        filiere: 'Cosmétiques & bien-être' },
  { match: ['détergent', 'detergent', 'nettoyage'],          filiere: 'Détergents & produits de nettoyage' },
  { match: ['peinture', 'vernis', 'colle', 'encre', 'adhés'], filiere: 'Peintures, colles & encres' },
  { match: ['pigment', 'colorant'],                           filiere: 'Pigments & colorants' },
  { match: ['polymère', 'polymere', 'résine', 'resine'],     filiere: 'Polymères & résines' },
  { match: ['verre'],                                         filiere: 'Verres' },
  { match: ['phytosanitaire', 'phyto'],                       filiere: 'Phytosanitaires' },
  { match: ['gaz industriel', 'gaz médical', 'gaz'],         filiere: 'Gaz industriels & médicaux' },
]

async function parseExcelFile(file: File): Promise<{ companies: Company[]; withSectors: number }> {
  const XLSX = await import('xlsx')
  const buffer = await file.arrayBuffer()
  const wb = XLSX.read(buffer, { type: 'array' })
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: null })

  if (rows.length < 2) return { companies: [], withSectors: 0 }

  // Find header row (first row that has "Code" or "Raison")
  let headerRowIdx = 0
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const r = rows[i] as unknown[]
    if (r.some(c => String(c ?? '').toLowerCase().includes('code') || String(c ?? '').toLowerCase().includes('raison'))) {
      headerRowIdx = i
      break
    }
  }

  const headers = (rows[headerRowIdx] as unknown[]).map(h => String(h ?? '').toLowerCase().trim())

  // Detect sector columns dynamically
  const sectorCols: { idx: number; filiere: string }[] = []
  headers.forEach((h, idx) => {
    if (!h) return
    for (const { match, filiere } of SECTOR_MAP) {
      if (match.some(m => h.includes(m))) {
        if (!sectorCols.some(sc => sc.idx === idx)) {
          sectorCols.push({ idx, filiere })
        }
        break
      }
    }
  })

  // Detect key columns (with fallback to known FCP Excel positions)
  const find = (terms: string[], fallback: number) => {
    const idx = headers.findIndex(h => terms.some(t => h.includes(t)))
    return idx >= 0 ? idx : fallback
  }
  const codeIdx     = find(['code firme', 'code'], 0)
  const nameIdx     = find(['raison sociale', 'raison'], 3)
  const cityIdx     = find(['ville'], 7)
  const activiteIdx = find(['activit'], 8)
  const effectifIdx = find(['effectif'], 9)
  const caIdx       = headers.findIndex(h => h === 'ca')
  const caExpIdx    = find(['ca_exp', 'ca exp'], 12)

  const companies: Company[] = []
  let withSectors = 0

  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i] as unknown[]
    if (!row) continue
    const code = String(row[codeIdx] ?? '').trim()
    const name = String(row[nameIdx] ?? '').trim()
    if (!code || !name) continue

    const sectorsChecked = Array.from(new Set(
      sectorCols
        .filter(({ idx }) => String(row[idx] ?? '').trim().toUpperCase() === 'X')
        .map(({ filiere }) => filiere)
    ))

    if (sectorsChecked.length > 0) withSectors++

    const caVal = caIdx >= 0 ? row[caIdx] : row[10]
    companies.push({
      code,
      name,
      city:     String(row[cityIdx] ?? '').trim(),
      ca:       caVal ? String(caVal) : undefined,
      caExp:    row[caExpIdx] ? String(row[caExpIdx]) : undefined,
      effectif: row[effectifIdx] ? String(row[effectifIdx]) : undefined,
      activite: row[activiteIdx] ? String(row[activiteIdx]).substring(0, 400) : undefined,
      sectorsChecked: sectorsChecked.length > 0 ? sectorsChecked : undefined,
    })
  }

  return { companies, withSectors }
}

function exportCSV(results: EnrichedResult[]) {
  const hdr = [
    'Code','Nom','Ville',
    'Type','Filière Principale','Secteur Réel',
    'CA','Tranche CA','CA Export','Part Export','Effectif',
    'Dirigeant','Email','Mobile','Site Web',
    'Filières Confirmées','Filières Rejetées','Filières Découvertes',
    'Confiance','Raisonnement','Sources',
  ]
  const rows = results.map(r => [
    r.code, `"${r.name}"`, r.city,
    r.type_entreprise||'', r.filiere_principale||'', r.secteur_reel||'',
    r.ca_montant||'', r.tranche_ca||'', r.ca_export||'', r.part_export||'', r.effectif_tranche||'',
    r.dirigeant||'', r.email_dirigeant||'', r.mobile_dirigeant||'', r.site_web||'',
    `"${(r.filieres_confirmees||[]).join(' | ')}"`,
    `"${(r.filieres_rejetees||[]).join(' | ')}"`,
    `"${(r.filieres_ajoutees||[]).join(' | ')}"`,
    `${Math.round((r.confiance||0)*100)}%`,
    `"${(r.raison||'').replace(/"/g,'""')}"`,
    `"${(r.sources||[]).join(', ')}"`,
  ].join(','))
  const blob = new Blob([[hdr.join(','), ...rows].join('\n')], { type: 'text/csv;charset=utf-8;' })
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob)
  a.download = `fcp_enrich_${new Date().toISOString().slice(0,10)}.csv`; a.click()
}

async function exportXLSX(results: EnrichedResult[]) {
  const XLSX = await import('xlsx')
  const ws_data = [
    ['Code','Nom','Ville','Type','Filière Principale','Secteur Réel',
     'CA','Tranche CA','CA Export','Part Export','Effectif',
     'Dirigeant','Email','Mobile','Site Web',
     'Filières Confirmées','Filières Rejetées','Filières Découvertes','Confiance','Raisonnement'],
    ...results.map(r => [
      r.code, r.name, r.city,
      r.type_entreprise||'', r.filiere_principale||'', r.secteur_reel||'',
      r.ca_montant||'', r.tranche_ca||'', r.ca_export||'', r.part_export||'', r.effectif_tranche||'',
      r.dirigeant||'', r.email_dirigeant||'', r.mobile_dirigeant||'', r.site_web||'',
      (r.filieres_confirmees||[]).join(' | '),
      (r.filieres_rejetees||[]).join(' | '),
      (r.filieres_ajoutees||[]).join(' | '),
      `${Math.round((r.confiance||0)*100)}%`, r.raison||'',
    ]),
  ]
  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.aoa_to_sheet(ws_data)
  ws['!cols'] = [10,30,16,22,26,22,18,16,16,12,10,22,26,16,30,30,30,30,10,50].map(w => ({ wch: w }))
  XLSX.utils.book_append_sheet(wb, ws, 'FCP Enrichissements')
  XLSX.writeFile(wb, `fcp_enrich_${new Date().toISOString().slice(0,10)}.xlsx`)
}

// ─── Sub-components ───────────────────────────────────────────────────────

function FiliereBadge({ filiere, small }: { filiere: string | null; small?: boolean }) {
  if (!filiere) return <span style={{ color: '#C4BEB8', fontSize: 12 }}>—</span>
  const col = FILIERE_COLORS[filiere] || { c: '#808080', bg: '#F0F0F0' }
  return (
    <span className="fcp-badge" style={{
      color: col.c, background: col.bg,
      fontSize: small ? 11 : 11.5,
      padding: small ? '2px 7px' : '3px 9px',
    }}>
      {filiere}
    </span>
  )
}

function TypeBadge({ type }: { type: string | null }) {
  if (!type) return <span style={{ color: '#C4BEB8', fontSize: 12 }}>—</span>
  const col = TYPE_COLORS[type] || { c: '#808080', bg: '#F0F0F0' }
  return (
    <span className="fcp-badge" style={{ color: col.c, background: col.bg }}>
      {type}
    </span>
  )
}

function ConfBar({ value }: { value: number }) {
  const pct = Math.round(value * 100)
  const color = pct >= 75 ? '#2A7A52' : pct >= 50 ? '#B87020' : '#B84040'
  return (
    <div className="fcp-conf-bar">
      <div className="fcp-conf-track">
        <div className="fcp-conf-fill" style={{ width: `${pct}%`, background: color }} />
      </div>
      <span style={{ fontSize: 12, color, fontWeight: 600, fontVariantNumeric: 'tabular-nums' }}>
        {pct}%
      </span>
    </div>
  )
}

// ── Sidebar ───────────────────────────────────────────────────────────────
function Sidebar({
  view, setView, progress, resultCount, keys, running,
}: {
  view: string; setView: (v: 'config' | 'results') => void
  progress: { done: number; total: number }
  resultCount: number; keys: { anthropic: string; tavily: string }
  running: boolean
}) {
  const pct = progress.total ? Math.round(progress.done / progress.total * 100) : 0
  return (
    <aside className="fcp-sidebar">
      <div className="fcp-logo">FCP<span>Intel</span></div>

      <nav className="fcp-nav">
        <button className={`fcp-nav-item ${view === 'config' ? 'active' : ''}`} onClick={() => setView('config')}>
          <Settings size={15} />
          Configuration
        </button>
        <button className={`fcp-nav-item ${view === 'results' ? 'active' : ''}`} onClick={() => setView('results')}>
          <Table2 size={15} />
          Résultats
          {resultCount > 0 && <span className="fcp-nav-badge">{resultCount}</span>}
        </button>
      </nav>

      {progress.total > 0 && (
        <div className="fcp-progress-card">
          <div className="fcp-progress-label">
            {running ? <span className="fcp-pulse">● En cours</span> : '● Traitement'}
          </div>
          <div className="fcp-progress-nums">
            {progress.done}<span style={{ color: 'rgba(255,255,255,0.3)', fontWeight: 400 }}>/{progress.total}</span>
          </div>
          <div className="fcp-progress-bar">
            <div className="fcp-progress-fill" style={{ width: `${pct}%` }} />
          </div>
          <div className="fcp-progress-sub">{pct}% complété</div>
        </div>
      )}

      <div className="fcp-status-section">
        <div className="fcp-status-row">
          <div className={`fcp-status-dot ${keys.anthropic ? 'on' : 'off'}`} />
          Anthropic {keys.anthropic ? 'connecté' : 'non configuré'}
        </div>
        <div className="fcp-status-row">
          <div className={`fcp-status-dot ${keys.tavily ? 'on' : 'off'}`} />
          Tavily {keys.tavily ? 'connecté' : 'non configuré'}
        </div>
      </div>
    </aside>
  )
}

// ── Config View ───────────────────────────────────────────────────────────
function ConfigView({
  anthropicKey, setAnthropicKey,
  tavilyKey, setTavilyKey,
  supabaseUrl, setSupabaseUrl,
  supabaseKey, setSupabaseKey,
  raw, setRaw,
  batchSize, setBatchSize,
  onStart, running, paused, onPause, onStop, onSyncDB,
  results,
  parsedFromFile, fileName, fileStats, onFileUpload, onClearFile,
  envConfig,
}: any) {
  const [showSupabase, setShowSupabase] = useState(false)
  const [inputMode, setInputMode] = useState<'file' | 'tsv'>(parsedFromFile ? 'file' : 'tsv')
  const fileInputRef = useRef<HTMLInputElement>(null)

  const companies = parsedFromFile ?? parseRaw(raw)
  const done = new Set((results as EnrichedResult[]).map(r => r.code))
  const remaining = companies.filter((c: Company) => !done.has(c.code)).length
  const withSectors = fileStats?.withSectors ?? companies.filter((c: Company) => c.sectorsChecked && c.sectorsChecked.length > 0).length

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault()
    const file = e.dataTransfer.files[0]
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) onFileUpload(file)
  }

  const EnvBadge = ({ label }: { label: string }) => (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '9px 13px', background: '#F0FBF5', border: '1.5px solid #B8E8CC', borderRadius: 8 }}>
      <span style={{ fontSize: 13, color: '#2A7A52' }}>🔒</span>
      <span style={{ fontSize: 13, color: '#2A7A52', fontWeight: 500 }}>{label} configurée via variable d'environnement</span>
    </div>
  )

  return (
    <div style={{ flex: 1, overflowY: 'auto', padding: '32px 36px', maxWidth: 720 }}>
      <div style={{ marginBottom: 28 }}>
        <h1 style={{ fontSize: 22, fontWeight: 700, color: '#1B1916', letterSpacing: '-0.5px', marginBottom: 6 }}>
          Configuration
        </h1>
        <p style={{ fontSize: 14, color: '#9E9994', margin: 0 }}>
          Clés API, données entreprises et paramètres d'enrichissement.
        </p>
      </div>

      {/* API Keys */}
      <div className="fcp-card">
        <div className="fcp-card-title">Clés API</div>
        <div className="fcp-field">
          <label className="fcp-label">Clé Anthropic *</label>
          {envConfig.anthropic
            ? <EnvBadge label="Anthropic" />
            : <input type="password" className="fcp-input" value={anthropicKey} onChange={e => setAnthropicKey(e.target.value)} placeholder="sk-ant-api03-..." />
          }
        </div>
        <div className="fcp-field">
          <label className="fcp-label">Clé Tavily *</label>
          {envConfig.tavily
            ? <EnvBadge label="Tavily" />
            : <input type="password" className="fcp-input" value={tavilyKey} onChange={e => setTavilyKey(e.target.value)} placeholder="tvly-..." />
          }
        </div>
        <details open={showSupabase} onToggle={e => setShowSupabase((e.target as HTMLDetailsElement).open)}>
          <summary style={{ fontSize: 12.5, color: '#9E9994', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, marginTop: 4, userSelect: 'none' }}>
            <Database size={13} />
            Supabase (base de données)
            {envConfig.supabase
              ? <span style={{ color: '#2A7A52', fontSize: 11, fontWeight: 600 }}>● Connectée via env</span>
              : showSupabase ? <ChevronDown size={13} /> : <ChevronRight size={13} />
            }
          </summary>
          <div style={{ marginTop: 14 }}>
            {envConfig.supabase ? (
              <>
                <EnvBadge label="Supabase" />
                <p style={{ fontSize: 12, color: '#9E9994', marginTop: 8 }}>Résultats sauvegardés et chargés automatiquement.</p>
              </>
            ) : (
              <>
                <div className="fcp-field">
                  <label className="fcp-label">URL Supabase</label>
                  <input type="text" className="fcp-input" value={supabaseUrl} onChange={e => setSupabaseUrl(e.target.value)} placeholder="https://xxx.supabase.co" />
                </div>
                <div className="fcp-field" style={{ marginBottom: 12 }}>
                  <label className="fcp-label">Clé Supabase (anon)</label>
                  <input type="password" className="fcp-input" value={supabaseKey} onChange={e => setSupabaseKey(e.target.value)} placeholder="eyJ..." />
                </div>
              </>
            )}
            <button onClick={onSyncDB} style={{ fontSize: 12.5, color: '#6B6560', background: '#F4F2EE', border: '1.5px solid #E8E5E0', borderRadius: 7, padding: '6px 14px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, fontFamily: 'inherit' }}>
              <RefreshCw size={12} /> Synchroniser depuis Supabase
            </button>
          </div>
        </details>
      </div>

      {/* Data — dual mode */}
      <div className="fcp-card">
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
          <div className="fcp-card-title" style={{ margin: 0 }}>Données entreprises</div>
          {/* Mode toggle */}
          <div style={{ display: 'flex', border: '1.5px solid #E8E5E0', borderRadius: 8, overflow: 'hidden' }}>
            {(['file', 'tsv'] as const).map(mode => (
              <button key={mode} onClick={() => setInputMode(mode)} style={{
                padding: '5px 13px', border: 'none', cursor: 'pointer', fontSize: 12, fontWeight: 500,
                fontFamily: 'inherit',
                background: inputMode === mode ? '#1B1916' : '#fff',
                color: inputMode === mode ? '#fff' : '#6B6560',
                transition: 'all 0.15s',
              }}>
                {mode === 'file' ? '📎 Fichier Excel' : '✏️ Paste TSV'}
              </button>
            ))}
          </div>
        </div>

        {inputMode === 'file' ? (
          <>
            {/* File drop zone */}
            <div
              onDrop={handleDrop}
              onDragOver={e => e.preventDefault()}
              onClick={() => fileInputRef.current?.click()}
              style={{
                border: `2px dashed ${parsedFromFile ? '#2A7A52' : '#D4CFC8'}`,
                borderRadius: 10,
                padding: '28px 24px',
                textAlign: 'center',
                cursor: 'pointer',
                background: parsedFromFile ? '#F4FBF7' : '#FAFAF9',
                transition: 'all 0.2s',
                marginBottom: 14,
              }}
            >
              <input
                ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }}
                onChange={e => { const f = e.target.files?.[0]; if (f) onFileUpload(f) }}
              />
              {parsedFromFile ? (
                <>
                  <div style={{ fontSize: 28, marginBottom: 8 }}>✅</div>
                  <div style={{ fontSize: 14, fontWeight: 600, color: '#1B1916', marginBottom: 4 }}>
                    {fileName}
                  </div>
                  <div style={{ fontSize: 13, color: '#2A7A52', marginBottom: 10 }}>
                    <strong>{companies.length}</strong> entreprises importées
                    {withSectors > 0 && <> · <strong>{withSectors}</strong> avec secteurs cochés ✓</>}
                  </div>
                  <button
                    onClick={e => { e.stopPropagation(); onClearFile() }}
                    style={{ fontSize: 12, color: '#B84040', background: 'none', border: '1px solid #E8E5E0', borderRadius: 6, padding: '4px 10px', cursor: 'pointer', fontFamily: 'inherit' }}
                  >
                    ✕ Retirer le fichier
                  </button>
                </>
              ) : (
                <>
                  <div style={{ fontSize: 32, marginBottom: 10 }}>📎</div>
                  <div style={{ fontSize: 14, fontWeight: 600, color: '#1B1916', marginBottom: 6 }}>
                    Glissez votre fichier Excel ici
                  </div>
                  <div style={{ fontSize: 13, color: '#9E9994', marginBottom: 12 }}>
                    ou cliquez pour parcourir
                  </div>
                  <div style={{ fontSize: 12, color: '#B8B2AA' }}>
                    .xlsx · .xls — Lit automatiquement les colonnes et les secteurs cochés (X)
                  </div>
                </>
              )}
            </div>

            {/* Columns read info */}
            {parsedFromFile && (
              <div className="fcp-hint">
                <strong>Colonnes lues automatiquement:</strong> Code Firme · Raison Sociale · Ville · CA · CA Export · Effectif · Activité · Secteurs cochés (X)<br />
                <strong>Les secteurs cochés seront vérifiés</strong> par Claude — confirmés ✓, rejetés ✗ ou complétés +
              </div>
            )}
          </>
        ) : (
          <>
            <div className="fcp-hint">
              <strong>Format TSV</strong> — copiez depuis Excel (colonnes A, D, H + optionnel):<br />
              <strong>Code</strong> · <strong>Raison Sociale</strong> · <strong>Ville</strong> [· CA · CA Export · Effectif · Activité]
            </div>
            <div className="fcp-field">
              <textarea className="fcp-input fcp-textarea" value={raw}
                onChange={e => setRaw(e.target.value)}
                placeholder={'MA0695900\tOCP S.A.\tCasablanca\nMA0436000\tBayer\tCasablanca\nMA0434700\tBasf Maroc\tCasablanca'}
              />
            </div>
          </>
        )}

        <div style={{ display: 'flex', gap: 16, fontSize: 12.5, color: '#9E9994' }}>
          <span><strong style={{ color: '#1B1916' }}>{companies.length}</strong> entreprises</span>
          <span><strong style={{ color: '#C8622A' }}>{remaining}</strong> à traiter</span>
          {done.size > 0 && <span><strong style={{ color: '#2A7A52' }}>{done.size}</strong> déjà enrichies</span>}
          {withSectors > 0 && <span><strong style={{ color: '#7040A0' }}>{withSectors}</strong> avec secteurs à vérifier</span>}
        </div>
      </div>

      {/* Settings */}
      <div className="fcp-card">
        <div className="fcp-card-title">Paramètres</div>
        <div className="fcp-field" style={{ marginBottom: 0 }}>
          <label className="fcp-label">Requêtes parallèles</label>
          <div className="fcp-batch-group">
            {[3, 5, 10].map(n => (
              <button key={n} className={`fcp-batch-btn ${batchSize === n ? 'active' : ''}`} onClick={() => setBatchSize(n)}>
                {n} parallèles
              </button>
            ))}
          </div>
          <p style={{ fontSize: 12, color: '#B8B2AA', marginTop: 8, marginBottom: 0 }}>
            Modèle: Claude Haiku (rapide & économique) · 3 recherches Tavily par entreprise
          </p>
        </div>
      </div>

      {/* Launch */}
      <div style={{ display: 'flex', gap: 10 }}>
        {!running ? (
          <button className="fcp-launch" onClick={onStart}
            disabled={(!anthropicKey && !envConfig.anthropic) || (!tavilyKey && !envConfig.tavily) || remaining === 0}>
            <Play size={16} fill="currentColor" />
            {remaining === 0 && companies.length > 0
              ? 'Tout déjà traité ✓'
              : `Lancer l'enrichissement — ${remaining} entreprises`}
          </button>
        ) : (
          <div style={{ display: 'flex', gap: 8, width: '100%' }}>
            <button className="fcp-launch" style={{ flex: 1, background: '#F4F2EE', color: '#1B1916', border: '1.5px solid #E8E5E0' }} disabled>
              <div className="fcp-spinner" style={{ borderColor: 'rgba(27,25,22,0.2)', borderTopColor: '#1B1916' }} />
              Traitement en cours…
            </button>
            <button className="fcp-ctrl-btn" onClick={onPause}>
              {paused ? <Play size={14} /> : <Pause size={14} />}
              {paused ? 'Reprendre' : 'Pause'}
            </button>
            <button className="fcp-ctrl-btn danger" onClick={onStop}>
              <Square size={14} /> Arrêter
            </button>
          </div>
        )}
      </div>
    </div>
  )
}

// ── Results View ──────────────────────────────────────────────────────────
function ResultsView({
  results, logs, running, paused, onPause, onStop, onSelect, selected,
}: {
  results: EnrichedResult[]; logs: string[]; running: boolean; paused: boolean
  onPause: () => void; onStop: () => void
  onSelect: (r: EnrichedResult | null) => void; selected: EnrichedResult | null
}) {
  const [search, setSearch] = useState('')
  const [typeFilter, setTypeFilter] = useState('all')
  const [filiereFilter, setFiliereFilter] = useState('all')
  const [logsOpen, setLogsOpen] = useState(true)
  const logsRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (logsRef.current) logsRef.current.scrollTop = logsRef.current.scrollHeight
  }, [logs])

  const filtered = results.filter(r => {
    const q = search.toLowerCase()
    const matchSearch = !q || r.name.toLowerCase().includes(q) ||
      r.city.toLowerCase().includes(q) ||
      (r.dirigeant || '').toLowerCase().includes(q) ||
      (r.filiere_principale || '').toLowerCase().includes(q)
    const matchType = typeFilter === 'all' || r.type_entreprise === typeFilter
    const matchFiliere = filiereFilter === 'all' || r.filiere_principale === filiereFilter
    return matchSearch && matchType && matchFiliere
  })

  const fabricants = results.filter(r => r.type_entreprise?.includes('Fabricant')).length
  const distribs = results.filter(r => r.type_entreprise?.includes('Distributeur') || r.type_entreprise === 'Négoce').length
  const imports = results.filter(r => r.type_entreprise?.includes('Importateur')).length
  const avgConf = results.length ? Math.round(results.reduce((s, r) => s + (r.confiance || 0), 0) / results.length * 100) : 0

  const uniqueTypes = Array.from(new Set(results.map(r => r.type_entreprise).filter(Boolean))) as string[]
  const uniqueFilieres = Array.from(new Set(results.map(r => r.filiere_principale).filter(Boolean))) as string[]

  if (results.length === 0 && !running) {
    return (
      <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
        <div className="fcp-empty">
          <div className="fcp-empty-icon">🧪</div>
          <div className="fcp-empty-title">Aucun résultat pour l'instant</div>
          <div className="fcp-empty-sub">Configurez vos clés API et collez vos données, puis lancez l'enrichissement.</div>
        </div>
      </div>
    )
  }

  return (
    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      {/* Top bar */}
      <div style={{
        padding: '14px 20px',
        background: '#fff',
        borderBottom: '1px solid #E8E5E0',
        display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap',
        flexShrink: 0,
      }}>
        <div className="fcp-search-wrap">
          <Search size={14} className="fcp-search-icon" />
          <input
            className="fcp-search"
            value={search}
            onChange={e => setSearch(e.target.value)}
            placeholder="Rechercher…"
          />
        </div>

        <div className="fcp-chips">
          <span style={{ fontSize: 11.5, color: '#B8B2AA', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px' }}>Type:</span>
          <button className={`fcp-chip ${typeFilter === 'all' ? 'active' : ''}`} onClick={() => setTypeFilter('all')}>Tous</button>
          {uniqueTypes.map(t => (
            <button key={t} className={`fcp-chip ${typeFilter === t ? 'active' : ''}`} onClick={() => setTypeFilter(t)}>
              {t}
            </button>
          ))}
        </div>

        {uniqueFilieres.length > 1 && (
          <div className="fcp-chips">
            <span style={{ fontSize: 11.5, color: '#B8B2AA', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.5px' }}>Filière:</span>
            <button className={`fcp-chip ${filiereFilter === 'all' ? 'active' : ''}`} onClick={() => setFiliereFilter('all')}>Toutes</button>
            {uniqueFilieres.map(f => (
              <button key={f} className={`fcp-chip ${filiereFilter === f ? 'active' : ''}`} onClick={() => setFiliereFilter(f)}>
                {f}
              </button>
            ))}
          </div>
        )}

        <div style={{ marginLeft: 'auto', display: 'flex', gap: 8, alignItems: 'center' }}>
          {running && (
            <>
              <button className="fcp-ctrl-btn" onClick={onPause}>
                {paused ? <Play size={13} /> : <Pause size={13} />}
                {paused ? 'Reprendre' : 'Pause'}
              </button>
              <button className="fcp-ctrl-btn danger" onClick={onStop}>
                <Square size={13} /> Arrêter
              </button>
            </>
          )}
          <button className="fcp-export-btn" onClick={() => exportCSV(results)}>
            <Download size={13} /> CSV
          </button>
          <button className="fcp-export-btn" onClick={() => exportXLSX(results)}>
            <Download size={13} /> Excel
          </button>
        </div>
      </div>

      {/* Stats */}
      <div style={{ padding: '12px 20px', flexShrink: 0 }}>
        <div className="fcp-stats">
          <div className="fcp-stat">
            <div className="fcp-stat-num">{results.length}</div>
            <div className="fcp-stat-label">Total</div>
          </div>
          <div className="fcp-stat">
            <div className="fcp-stat-num" style={{ color: '#2A7A52' }}>{fabricants}</div>
            <div className="fcp-stat-label">Fabricants</div>
          </div>
          <div className="fcp-stat">
            <div className="fcp-stat-num" style={{ color: '#2850A0' }}>{distribs}</div>
            <div className="fcp-stat-label">Distributeurs</div>
          </div>
          <div className="fcp-stat">
            <div className="fcp-stat-num" style={{ color: '#7040A0' }}>{imports}</div>
            <div className="fcp-stat-label">Importateurs</div>
          </div>
          <div className="fcp-stat">
            <div className="fcp-stat-num" style={{ color: avgConf >= 70 ? '#2A7A52' : '#B87020' }}>{avgConf}%</div>
            <div className="fcp-stat-label">Conf. moy.</div>
          </div>
          <div className="fcp-stat">
            <div className="fcp-stat-num">{filtered.length}</div>
            <div className="fcp-stat-label">Affichés</div>
          </div>
        </div>
      </div>

      {/* Table */}
      <div className="fcp-table-wrap" style={{ padding: '0 20px', flex: 1 }}>
        <table className="fcp-table">
          <thead>
            <tr>
              <th style={{ minWidth: 200 }}>Raison Sociale</th>
              <th>Ville</th>
              <th>Type</th>
              <th style={{ minWidth: 220 }}>Filières confirmées</th>
              <th>Dirigeant</th>
              <th style={{ minWidth: 140 }}>CA</th>
              <th>Tranche CA</th>
              <th style={{ minWidth: 130 }}>CA Export</th>
              <th>Part Export</th>
              <th>Effectif</th>
              <th>Confiance</th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={9} style={{ textAlign: 'center', color: '#C4BEB8', padding: '32px' }}>Aucun résultat correspondant</td></tr>
            )}
            {filtered.map(r => (
              <tr
                key={r.code}
                onClick={() => onSelect(selected?.code === r.code ? null : r)}
                className={selected?.code === r.code ? 'selected' : ''}
              >
                <td style={{ fontWeight: 500 }}>
                  {r.error && <AlertCircle size={12} style={{ color: '#B84040', marginRight: 6, display: 'inline' }} />}
                  {r.name}
                </td>
                <td style={{ color: '#6B6560' }}>{r.city}</td>
                <td><TypeBadge type={r.type_entreprise || null} /></td>
                <td style={{ maxWidth: 280, whiteSpace: 'normal', verticalAlign: 'middle' }}>
                  {(() => {
                    const filieres = r.filieres_confirmees?.length > 0
                      ? r.filieres_confirmees
                      : r.filiere_principale ? [r.filiere_principale] : []
                    const isHors = filieres.length === 1 && filieres[0] === 'Hors secteur FCP'
                    return filieres.length > 0
                      ? <div style={{ display: 'flex', flexWrap: 'wrap', gap: 3, alignItems: 'center' }}>
                          {filieres.map((f: string) => <FiliereBadge key={f} filiere={f} small />)}
                          {isHors && r.secteur_reel && (
                            <span style={{ fontSize: 11.5, color: '#9E9994', fontStyle: 'italic' }}>
                              {r.secteur_reel}
                            </span>
                          )}
                        </div>
                      : <span style={{ color: '#C4BEB8', fontSize: 12 }}>—</span>
                  })()}
                </td>
                <td style={{ color: r.dirigeant ? '#1B1916' : '#C4BEB8' }}>
                  {r.dirigeant || '—'}
                </td>
                <td style={{ fontFamily: "'DM Mono', monospace", fontSize: 12, color: r.ca_montant ? '#1B1916' : '#C4BEB8', maxWidth: 160 }}>
                  {r.ca_montant || '—'}
                </td>
                <td style={{ fontFamily: "'DM Mono', monospace", fontSize: 12.5, color: r.tranche_ca ? '#1B1916' : '#C4BEB8' }}>
                  {r.tranche_ca || '—'}
                </td>
                <td style={{ fontFamily: "'DM Mono', monospace", fontSize: 12, color: r.ca_export ? '#2A7A52' : '#C4BEB8' }}>
                  {r.ca_export || '—'}
                </td>
                <td style={{ fontFamily: "'DM Mono', monospace", fontSize: 12, color: r.part_export ? '#2A7A52' : '#C4BEB8', fontWeight: r.part_export ? 600 : 400 }}>
                  {r.part_export || '—'}
                </td>
                <td style={{ color: r.effectif_tranche ? '#1B1916' : '#C4BEB8' }}>
                  {r.effectif_tranche || '—'}
                </td>
                <td><ConfBar value={r.confiance || 0} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Logs Drawer */}
      <div className="fcp-logs-drawer">
        <div className="fcp-logs-header" onClick={() => setLogsOpen(o => !o)}>
          {logsOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
          Journal d'activité
          {running && <div className="fcp-spinner" style={{ marginLeft: 8, width: 12, height: 12, borderWidth: 1.5 }} />}
          <span style={{ marginLeft: 'auto', fontWeight: 400, fontSize: 11, color: '#C4BEB8' }}>
            {logs.length} entrées
          </span>
        </div>
        {logsOpen && (
          <div className="fcp-logs-body" ref={logsRef}>
            {logs.map((l, i) => {
              const cls = l.includes('✅') || l.includes('💾') ? 'ok'
                : l.includes('❌') ? 'err'
                : l.includes('⚡') || l.includes('🚀') ? 'head'
                : ''
              return <div key={i} className={`fcp-log-line ${cls}`}>{l}</div>
            })}
          </div>
        )}
      </div>
    </div>
  )
}

// ── Detail Panel ──────────────────────────────────────────────────────────
function DetailPanel({ result, onClose }: { result: EnrichedResult | null; onClose: () => void }) {
  const open = result !== null
  const r = result

  return (
    <>
      <div className={`fcp-backdrop ${open ? 'open' : ''}`} onClick={onClose} />
      <aside className={`fcp-panel ${open ? 'open' : ''}`}>
        {r && (
          <>
            <div className="fcp-panel-header">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                <div style={{ flex: 1, marginRight: 12 }}>
                  <div className="fcp-panel-name">{r.name}</div>
                  <div className="fcp-panel-code">{r.code} · {r.city}</div>
                  <div style={{ display: 'flex', gap: 7, flexWrap: 'wrap' }}>
                    <TypeBadge type={r.type_entreprise || null} />
                    {r.filiere_principale && <FiliereBadge filiere={r.filiere_principale} small />}
                    <span className="fcp-badge" style={{
                      color: r.confiance >= 0.75 ? '#2A7A52' : r.confiance >= 0.5 ? '#B87020' : '#B84040',
                      background: r.confiance >= 0.75 ? '#E0F0E8' : r.confiance >= 0.5 ? '#F5EDDC' : '#F5E5E5',
                    }}>
                      {Math.round(r.confiance * 100)}% confiance
                    </span>
                  </div>
                </div>
                <button
                  onClick={onClose}
                  style={{
                    border: 'none', background: '#F4F2EE', borderRadius: 8,
                    width: 32, height: 32, display: 'flex', alignItems: 'center',
                    justifyContent: 'center', cursor: 'pointer', color: '#9E9994', flexShrink: 0,
                  }}
                >
                  <X size={15} />
                </button>
              </div>
            </div>

            <div className="fcp-panel-body">

              {/* Direction */}
              <div className="fcp-panel-section">
                <div className="fcp-section-label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <User size={12} /> Direction
                </div>
                <Row label="Dirigeant" value={r.dirigeant} />
                <Row label="Email" value={r.email_dirigeant} link={r.email_dirigeant ? `mailto:${r.email_dirigeant}` : undefined} mono />
                <Row label="Mobile" value={r.mobile_dirigeant} mono />
                <Row label="Site web" value={r.site_web} link={r.site_web || undefined} mono short />
              </div>

              {/* Finances */}
              <div className="fcp-panel-section">
                <div className="fcp-section-label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <BarChart3 size={12} /> Données financières
                </div>
                <Row label="CA" value={r.ca_montant} mono />
                <Row label="Tranche CA" value={r.tranche_ca} mono />
                <Row label="CA Export" value={r.ca_export} mono />
                <Row label="Part export" value={r.part_export} mono highlight={r.part_export ? '#2A7A52' : undefined} />
                <Row label="Effectif" value={r.effectif_tranche} />
              </div>

              {/* Filières */}
              <div className="fcp-panel-section">
                <div className="fcp-section-label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <Layers size={12} /> Filières d'activité
                </div>

                {/* Hors FCP → show real sector */}
                {r.filiere_principale === 'Hors secteur FCP' && r.secteur_reel && (
                  <div style={{
                    background: '#F9F8F6', border: '1px solid #E8E5E0',
                    borderRadius: 8, padding: '8px 12px', marginBottom: 12,
                    display: 'flex', alignItems: 'center', gap: 8,
                  }}>
                    <span style={{ fontSize: 12, color: '#9E9994' }}>Secteur réel :</span>
                    <span style={{ fontSize: 13, fontWeight: 600, color: '#1B1916' }}>{r.secteur_reel}</span>
                  </div>
                )}

                {/* ── Checked sectors from Excel — show ALL with verdict ── */}
                {(r.sectorsChecked && r.sectorsChecked.length > 0) ? (
                  <>
                    <div style={{ fontSize: 11, color: '#B8B2AA', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.6px', marginBottom: 8, marginTop: 2 }}>
                      Cochés dans la base ({r.sectorsChecked.length})
                    </div>
                    {r.sectorsChecked.map((f: string) => {
                      const isConfirmed = (r.filieres_confirmees || []).includes(f)
                      const isRejected  = (r.filieres_rejetees  || []).includes(f)
                      const status = isConfirmed ? 'ok' : isRejected ? 'no' : 'no'
                      return (
                        <div key={f} className="fcp-filiere-row">
                          <div className={`fcp-filiere-icon ${status}`}>
                            {isConfirmed ? <Check size={10} /> : <Minus size={10} />}
                          </div>
                          <span className="fcp-filiere-name" style={{ textDecoration: (!isConfirmed) ? 'line-through' : 'none', opacity: (!isConfirmed) ? 0.55 : 1 }}>{f}</span>
                          <span className={`fcp-filiere-status ${status}`}>
                            {isConfirmed ? 'Confirmé ✓' : 'Non confirmé ✗'}
                          </span>
                        </div>
                      )
                    })}

                    {/* Discovered sectors (not in original checked list) */}
                    {(r.filieres_ajoutees || []).length > 0 && (
                      <>
                        <div style={{ fontSize: 11, color: '#B8B2AA', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.6px', marginBottom: 8, marginTop: 12 }}>
                          Découverts par la recherche
                        </div>
                        {(r.filieres_ajoutees || []).map((f: string) => (
                          <div key={f} className="fcp-filiere-row">
                            <div className="fcp-filiere-icon new"><Plus size={10} /></div>
                            <span className="fcp-filiere-name">{f}</span>
                            <span className="fcp-filiere-status new">Découvert +</span>
                          </div>
                        ))}
                      </>
                    )}
                  </>
                ) : (
                  /* No checked sectors — show what Claude found */
                  <>
                    {(r.filieres_confirmees || []).map((f: string) => (
                      <div key={f} className="fcp-filiere-row">
                        <div className="fcp-filiere-icon ok"><Check size={10} /></div>
                        <span className="fcp-filiere-name">{f}</span>
                        <span className="fcp-filiere-status ok">Identifié</span>
                      </div>
                    ))}
                    {(r.filieres_ajoutees || []).map((f: string) => (
                      <div key={f} className="fcp-filiere-row">
                        <div className="fcp-filiere-icon new"><Plus size={10} /></div>
                        <span className="fcp-filiere-name">{f}</span>
                        <span className="fcp-filiere-status new">Découvert</span>
                      </div>
                    ))}
                    {(r.filieres_rejetees || []).map((f: string) => (
                      <div key={f} className="fcp-filiere-row">
                        <div className="fcp-filiere-icon no"><Minus size={10} /></div>
                        <span className="fcp-filiere-name" style={{ textDecoration: 'line-through', opacity: 0.55 }}>{f}</span>
                        <span className="fcp-filiere-status no">Rejeté</span>
                      </div>
                    ))}
                    {(!r.filieres_confirmees?.length && !r.filieres_ajoutees?.length) && (
                      <span style={{ fontSize: 13, color: '#C4BEB8' }}>Aucune filière déterminée</span>
                    )}
                  </>
                )}
              </div>

              {/* Intelligence */}
              <div className="fcp-panel-section">
                <div className="fcp-section-label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <TrendingUp size={12} /> Raisonnement
                </div>
                {r.raison ? (
                  <div className="fcp-raison">{r.raison}</div>
                ) : (
                  <span style={{ fontSize: 13, color: '#C4BEB8' }}>—</span>
                )}
                {r.error && (
                  <div style={{
                    background: '#F5E5E5', border: '1px solid #EAC0C0',
                    borderRadius: 8, padding: '8px 12px', fontSize: 12.5, color: '#B84040',
                  }}>
                    Erreur: {r.error}
                  </div>
                )}
              </div>

              {/* Sources */}
              {(r.sources || []).length > 0 && (
                <div className="fcp-panel-section">
                  <div className="fcp-section-label" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <ExternalLink size={12} /> Sources
                  </div>
                  <div className="fcp-sources">
                    {(r.sources || []).map((src, i) => (
                      <a key={i} href={src} target="_blank" rel="noopener noreferrer" className="fcp-source-link">
                        {src}
                      </a>
                    ))}
                  </div>
                </div>
              )}

              {/* Processed at */}
              <div style={{ padding: '12px 22px', fontSize: 11.5, color: '#C4BEB8', fontFamily: "'DM Mono', monospace" }}>
                Traité le {new Date(r.processed_at).toLocaleString('fr-FR')}
              </div>
            </div>
          </>
        )}
      </aside>
    </>
  )
}

// Small helper row component
function Row({ label, value, link, mono, highlight, short }: {
  label: string; value: string | null | undefined; link?: string
  mono?: boolean; highlight?: string; short?: boolean
}) {
  return (
    <div className="fcp-field-row">
      <span className="fcp-field-name">{label}</span>
      {link && value ? (
        <a href={link} target="_blank" rel="noopener noreferrer"
          className="fcp-field-val link"
          style={{ fontFamily: mono ? "'DM Mono', monospace" : undefined, fontSize: mono ? 12 : undefined }}>
          {short && value.length > 32 ? value.substring(0, 32) + '…' : value}
        </a>
      ) : (
        <span className={`fcp-field-val ${!value ? 'nil' : ''} ${mono ? 'mono' : ''}`}
          style={{ color: (value && highlight) ? highlight : undefined }}>
          {value || '—'}
        </span>
      )}
    </div>
  )
}

// ─── Main Page ────────────────────────────────────────────────────────────
export default function Home() {
  // Keys
  const [anthropicKey, setAnthropicKeyRaw] = useState('')
  const [tavilyKey, setTavilyKeyRaw] = useState('')
  const [supabaseUrl, setSupabaseUrlRaw] = useState('')
  const [supabaseKey, setSupabaseKeyRaw] = useState('')

  // Data
  const [raw, setRawRaw] = useState('')
  const [parsedFromFile, setParsedFromFile] = useState<Company[] | null>(null)
  const [fileName, setFileName] = useState('')
  const [fileStats, setFileStats] = useState<{ total: number; withSectors: number } | null>(null)
  const [results, setResultsRaw] = useState<EnrichedResult[]>([])
  const [logs, setLogs] = useState<string[]>(['Bienvenue sur FCP Intel — Agent prêt.'])

  // UI
  const [view, setView] = useState<'config' | 'results'>('config')
  const [selected, setSelected] = useState<EnrichedResult | null>(null)
  const [batchSize, setBatchSizeRaw] = useState(5)
  const [envConfig, setEnvConfig] = useState({ anthropic: false, tavily: false, supabase: false })

  // Run
  const [running, setRunning] = useState(false)
  const [paused, setPaused] = useState(false)
  const [progress, setProgress] = useState({ done: 0, total: 0 })
  const pauseRef = useRef(false)
  const stopRef = useRef(false)
  const [hydrated, setHydrated] = useState(false)

  // Hydrate from localStorage
  useEffect(() => {
    setAnthropicKeyRaw(load('fcp_ak', ''))
    setTavilyKeyRaw(load('fcp_tk', ''))
    setSupabaseUrlRaw(load('fcp_su', ''))
    setSupabaseKeyRaw(load('fcp_sk', ''))
    setRawRaw(load('fcp_raw', ''))
    setResultsRaw(load('fcp_results', []))
    setBatchSizeRaw(load('fcp_batch', 5))
    setHydrated(true)
  }, [])

  // Health check + auto-sync from Supabase when env vars are configured
  useEffect(() => {
    if (!hydrated) return
    fetch('/api/health')
      .then(r => r.json())
      .then(cfg => {
        setEnvConfig(cfg)
        if (cfg.supabase) {
          addLog('🔗 Supabase détecté — synchronisation automatique…')
          fetch('/api/results', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({}) })
            .then(r => r.json())
            .then(data => {
              if (data.results?.length) {
                setResults(() => data.results)
                addLog(`💾 ${data.results.length} résultats chargés depuis Supabase`)
                setView('results')
              }
            })
            .catch(() => {})
        }
      })
      .catch(() => {})
  }, [hydrated])

  const setAnthropicKey = (v: string) => { setAnthropicKeyRaw(v); save('fcp_ak', v) }
  const setTavilyKey    = (v: string) => { setTavilyKeyRaw(v);    save('fcp_tk', v) }
  const setSupabaseUrl  = (v: string) => { setSupabaseUrlRaw(v);  save('fcp_su', v) }
  const setSupabaseKey  = (v: string) => { setSupabaseKeyRaw(v);  save('fcp_sk', v) }
  const setRaw          = (v: string) => { setRawRaw(v);          save('fcp_raw', v) }
  const setBatchSize    = (v: number) => { setBatchSizeRaw(v);    save('fcp_batch', v) }
  const setResults      = (fn: (p: EnrichedResult[]) => EnrichedResult[]) =>
    setResultsRaw(prev => { const next = fn(prev); save('fcp_results', next); return next })

  const addLog  = useCallback((m: string) => setLogs(p => [...p.slice(-500), m]), [])
  const addLogs = useCallback((ms: string[]) => setLogs(p => [...p.slice(-500), ...ms]), [])

  const handleFileUpload = async (file: File) => {
    try {
      addLog(`📎 Lecture du fichier: ${file.name}…`)
      const { companies, withSectors } = await parseExcelFile(file)
      setParsedFromFile(companies)
      setFileName(file.name)
      setFileStats({ total: companies.length, withSectors })
      addLog(`✅ ${companies.length} entreprises importées · ${withSectors} avec secteurs cochés`)
    } catch (e: any) {
      addLog(`❌ Erreur lecture fichier: ${e.message}`)
    }
  }

  const handleClearFile = () => {
    setParsedFromFile(null)
    setFileName('')
    setFileStats(null)
  }

  const start = async () => {
    if (!anthropicKey || !tavilyKey) { addLog('❌ Clés Anthropic/Tavily manquantes'); return }
    const companies = parsedFromFile ?? parseRaw(raw)
    const done = new Set(results.map(r => r.code))
    const todo = companies.filter(c => !done.has(c.code))
    if (!todo.length) { addLog('✅ Toutes les entreprises sont déjà enrichies'); return }

    const withSec = todo.filter(c => c.sectorsChecked && c.sectorsChecked.length > 0).length

    setView('results')
    setRunning(true); setPaused(false)
    stopRef.current = false; pauseRef.current = false
    setProgress({ done: 0, total: todo.length })
    setLogs([])
    addLog(`🚀 Démarrage — ${todo.length} entreprises — ${batchSize} parallèles${withSec > 0 ? ` · ${withSec} avec secteurs à vérifier` : ''}`)

    const totalWaves = Math.ceil(todo.length / batchSize)
    for (let w = 0; w < totalWaves; w++) {
      if (stopRef.current) break
      while (pauseRef.current) await sleep(300)

      const batch = todo.slice(w * batchSize, (w + 1) * batchSize)
      addLog(`\n⚡ Vague ${w + 1}/${totalWaves} — ${batch.map(c => c.name).join(' · ')}`)

      try {
        const res = await fetch('/api/enrich', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            companies: batch,
            anthropicKey, tavilyKey, supabaseUrl, supabaseKey,
            concurrency: batchSize,
          }),
        })
        const data = await res.json()
        if (data.logs) addLogs(data.logs)
        if (data.results) {
          setResults(prev => [...prev, ...data.results])
          setProgress(p => ({ ...p, done: Math.min(p.done + batch.length, todo.length) }))
        }
      } catch (e: any) {
        addLog(`❌ Vague ${w + 1}: ${e.message}`)
      }

      if (w < totalWaves - 1 && !stopRef.current) await sleep(400)
    }
    setRunning(false)
    addLog(`\n✅ Terminé — ${todo.length} entreprises traitées`)
  }

  const onPause = () => {
    pauseRef.current = !pauseRef.current
    setPaused(pauseRef.current)
    addLog(pauseRef.current ? '⏸ En pause' : '▶ Reprise')
  }
  const onStop = () => {
    stopRef.current = true; pauseRef.current = false
    setPaused(false); setRunning(false)
    addLog('⏹ Arrêté')
  }
  const onSyncDB = async () => {
    if (!supabaseUrl || !supabaseKey) { addLog('⚠ Supabase non configuré'); return }
    addLog('🔄 Synchronisation depuis Supabase…')
    try {
      const res = await fetch('/api/results', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ supabaseUrl, supabaseKey }),
      })
      const data = await res.json()
      if (data.results?.length) {
        setResults(() => data.results)
        addLog(`💾 ${data.results.length} résultats chargés depuis la DB`)
        setView('results')
      } else addLog('ℹ DB vide ou aucun résultat')
    } catch (e: any) { addLog(`❌ DB sync: ${e.message}`) }
  }

  if (!hydrated) return null

  return (
    <div style={{ display: 'flex', height: '100vh', overflow: 'hidden', fontFamily: "'DM Sans', system-ui, sans-serif" }}>
      <Sidebar
        view={view} setView={setView}
        progress={progress}
        resultCount={results.length}
        keys={{ anthropic: anthropicKey, tavily: tavilyKey }}
        running={running}
      />
      <main style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', background: '#F4F2EE' }}>
        {view === 'config' ? (
          <ConfigView
            anthropicKey={anthropicKey} setAnthropicKey={setAnthropicKey}
            tavilyKey={tavilyKey} setTavilyKey={setTavilyKey}
            supabaseUrl={supabaseUrl} setSupabaseUrl={setSupabaseUrl}
            supabaseKey={supabaseKey} setSupabaseKey={setSupabaseKey}
            raw={raw} setRaw={setRaw}
            batchSize={batchSize} setBatchSize={setBatchSize}
            onStart={start} running={running} paused={paused}
            onPause={onPause} onStop={onStop} onSyncDB={onSyncDB}
            results={results}
            parsedFromFile={parsedFromFile}
            fileName={fileName}
            fileStats={fileStats}
            onFileUpload={handleFileUpload}
            onClearFile={handleClearFile}
            envConfig={envConfig}
          />
        ) : (
          <ResultsView
            results={results} logs={logs}
            running={running} paused={paused}
            onPause={onPause} onStop={onStop}
            onSelect={setSelected} selected={selected}
          />
        )}
      </main>
      <DetailPanel result={selected} onClose={() => setSelected(null)} />
    </div>
  )
}
