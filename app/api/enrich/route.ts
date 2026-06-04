import { NextRequest, NextResponse } from 'next/server'
import Anthropic from '@anthropic-ai/sdk'
import { createClient } from '@supabase/supabase-js'

// ── FCP Sectors reference ──────────────────────────────────────────────────
const FCP_SECTORS: Record<string, string[]> = {
  "Engrais & fertilisants": [
    "Engrais azotés (urée, ammonitrates, solutions azotées)",
    "Engrais phosphatés", "Engrais potassiques", "Engrais NPK",
    "Engrais liquides & ferti-irrigation", "Engrais organiques & organo-minéraux",
    "Amendements du sol",
  ],
  "Produits de commodité": [
    "Acides inorganiques (sulfurique, chlorhydrique, nitrique)",
    "Bases et alcalis (soude, potasse, ammoniac)",
    "Hypochlorites de soude (eau de Javel)",
    "Solvants organiques", "Sels techniques", "Oxydants & réducteurs",
    "Charges minérales (carbonates, silices, talc)",
  ],
  "Traitement de l'eau": [
    "Coagulants & floculants",
    "Désinfectants (chlore, hypochlorite, dioxyde de chlore, ozone)",
    "Correcteurs de pH", "Produits pour eau potable",
    "Traitement des eaux usées & des boues",
  ],
  "Cosmétiques & bien-être": [
    "Huiles essentielles", "Parfums & arômes", "Soins visage", "Soins corps",
    "Soins capillaires", "Hygiène (gels douche, savons, déodorants)",
    "Maquillage", "Parfums",
  ],
  "Détergents & produits de nettoyage": [
    "Lessives (poudre, liquide, capsules)", "Assouplissants textiles",
    "Détergents vaisselle", "Nettoyants sols & surfaces",
    "Désinfectants & produits d'hygiène professionnelle", "Nettoyants sanitaires",
    "Détergents industriels & agroalimentaires",
    "MP & Ingrédients pour détergence (tensioactifs, SLES)",
  ],
  "Peintures, colles & encres": [
    "Peintures bâtiment (intérieur/extérieur)", "Peintures industrielles & anticorrosion",
    "Peintures pour bois & métaux", "Étanchéité", "Vernis & laques décoratives",
    "Revêtements de sols & résines de protection",
    "Colles bâtiment", "Colles pour bois & papier", "Adhésifs industriels",
    "Colles pour emballage & étiquetage", "Encres offset, flexo, hélio",
    "Encres pour impression numérique", "Vernis de surfaçage & de protection",
  ],
  "Pigments & colorants": [
    "Pigments minéraux (oxydes, dioxyde de titane)",
    "Pigments organiques", "Pigments à effets (nacrés, métalliques)",
    "Colorants pour plastiques", "Colorants pour peintures & encres",
    "Colorants textiles", "Colorants alimentaires & boissons",
    "Colorants pour cosmétiques",
  ],
  "Polymères & résines": [
    "Polymères thermoplastiques (PE, PP, PVC, PET, PS)",
    "Polymères techniques (PA, POM, PC, PBT)",
    "Résines thermodurcissables (époxy, polyester, vinylester)",
    "Résines acryliques & alkydes", "Élastomères & caoutchoucs",
    "Résines pour composites & stratifiés", "Masterbatches & compounds formulés",
  ],
  "Verres": [
    "Verre pour emballage (bouteilles, bocaux)",
    "Verres de table et verres décorés",
  ],
  "Phytosanitaires": [
    "Herbicides", "Fongicides", "Insecticides & acaricides", "Nématicides",
    "Régulateurs de croissance", "Produits de biocontrôle",
  ],
  "Gaz industriels & médicaux": [
    "Gaz de l'air (oxygène, azote)",
    "Gaz combustibles (acétylène, hydrogène)",
    "Gaz de protection pour soudage & métallurgie",
    "Gaz réfrigérants", "Gaz médicaux", "Gaz pour laboratoires & analyses",
  ],
}

const SECTOR_NAMES = Object.keys(FCP_SECTORS)

// ── Helpers ───────────────────────────────────────────────────────────────
async function tavilySearch(query: string, key: string): Promise<string> {
  const res = await fetch('https://api.tavily.com/search', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      api_key: key,
      query,
      max_results: 5,
      include_answer: true,
      search_depth: 'basic',
    }),
  })
  const data = await res.json()
  if (!res.ok) throw new Error(`Tavily error: ${data.message || res.statusText}`)
  const snippets = (data.results || []).map(
    (r: { url: string; title: string; content: string }) =>
      `[${r.title}]\n${r.url}\n${r.content.substring(0, 600)}`
  ).join('\n\n---\n\n')
  return (data.answer ? `RÉSUMÉ TAVILY: ${data.answer}\n\n` : '') + snippets
}

// ── Main enrichment function ───────────────────────────────────────────────
type CompanyInput = {
  code: string
  name: string
  city: string
  ca?: string
  caExp?: string
  effectif?: string
  activite?: string
  sectorsChecked?: string[]   // Filières cochées dans l'Excel (à vérifier)
}

// ── CA formatting helpers ──────────────────────────────────────────────────
function formatCANum(val: string | undefined): string | undefined {
  if (!val) return undefined
  const n = parseFloat(String(val).replace(/[^0-9.]/g, ''))
  if (isNaN(n) || n === 0) return String(val).trim() || undefined
  if (n >= 1_000_000_000) return `${(n / 1_000_000_000).toFixed(2)} Md MAD`
  if (n >= 1_000_000)     return `${Math.round(n / 1_000_000)} M MAD`
  if (n >= 1_000)         return `${Math.round(n / 1_000)} K MAD`
  return `${n} MAD`
}

function deriveTrancheCA(val: string | undefined): string | undefined {
  if (!val) return undefined
  const n = parseFloat(String(val).replace(/[^0-9.]/g, ''))
  if (isNaN(n) || n === 0) return undefined
  if (n >= 500_000_000) return '> 500M MAD'
  if (n >= 100_000_000) return '100–500M MAD'
  if (n >= 50_000_000)  return '50–100M MAD'
  if (n >= 10_000_000)  return '10–50M MAD'
  return '< 10M MAD'
}

async function enrichOne(
  company: CompanyInput,
  anthropicKey: string,
  tavilyKey: string,
  supabaseUrl?: string,
  supabaseKey?: string,
) {
  const logs: string[] = []
  const log = (m: string) => logs.push(`[${new Date().toLocaleTimeString('fr-FR')}] ${m}`)

  log(`▶ ${company.name} (${company.city})`)

  // ── Pre-enrich from Excel data ───────────────────────────────────────────
  const caFormatted      = formatCANum(company.ca)
  const caExpFormatted   = formatCANum(company.caExp)
  const trancheFromExcel = deriveTrancheCA(company.ca)
  const hasFinancials    = !!(caFormatted || caExpFormatted)

  // ── 3 targeted Tavily searches ──────────────────────────────────────────
  log(`  🔍 Direction & contact`)
  const s1 = await tavilySearch(
    `"${company.name}" Maroc PDG "directeur général" dirigeant contact`,
    tavilyKey
  )

  // If we already have good financial data from Excel, search for export/sector details instead
  const financialQuery = hasFinancials
    ? `"${company.name}" Maroc export international part chiffre affaires effectif employés`
    : `"${company.name}" Maroc "chiffre d'affaires" millions milliards MAD 2022 2023 2024 export`

  log(`  🔍 Finances${hasFinancials ? ' (CA connu, focus export)' : ' (CA inconnu, recherche complète)'}`)
  const s2 = await tavilySearch(financialQuery, tavilyKey)

  log(`  🔍 Activité & produits`)
  const s3 = await tavilySearch(
    `"${company.name}" Maroc produits activité secteur fournisseur fabricant distributeur`,
    tavilyKey
  )

  log(`  ✓ Recherches terminées`)

  // ── Build context block ──────────────────────────────────────────────────
  const knownParts: string[] = []
  if (caFormatted)    knownParts.push(`CA Excel (source directe): ${caFormatted}`)
  if (caExpFormatted) knownParts.push(`CA Export Excel (source directe): ${caExpFormatted}`)
  if (company.effectif) knownParts.push(`Effectif Excel: ${company.effectif}`)
  if (company.activite) knownParts.push(`Activité déclarée: ${company.activite.substring(0, 350)}`)

  const hasSectors = company.sectorsChecked && company.sectorsChecked.length > 0
  const sectorsBlock = hasSectors
    ? `\nSECTEURS COCHÉS DANS LA BASE (fiabilité INCERTAINE — à vérifier rigoureusement):
${company.sectorsChecked!.map(s => `  • ${s}`).join('\n')}
⚠️ Ces secteurs ont été saisis manuellement et peuvent être incorrects. Vérifie chacun avec les sources.\n`
    : ''

  const sectorsRef = SECTOR_NAMES.map((s, i) =>
    `${i + 1}. ${s}: ${FCP_SECTORS[s].slice(0, 4).join(' / ')}${FCP_SECTORS[s].length > 4 ? '…' : ''}`
  ).join('\n')

  // ── Claude prompt ────────────────────────────────────────────────────────
  const prompt = `Expert en intelligence économique — entreprises marocaines chimie/parachimie (FCP).

ENTREPRISE: "${company.name}" | Ville: ${company.city} | Code: ${company.code}
${knownParts.length ? `\nDONNÉES CONNUES (priorité absolue — utilise-les directement):\n${knownParts.join('\n')}` : ''}${sectorsBlock}
=== RECHERCHE 1 — Direction & Contact ===
${s1.substring(0, 1600)}

=== RECHERCHE 2 — Finances & Effectif ===
${s2.substring(0, 1600)}

=== RECHERCHE 3 — Activité & Secteur ===
${s3.substring(0, 1600)}

=== 11 FILIÈRES FCP ===
${sectorsRef}

─────────────────────────────────────────────
INSTRUCTIONS:

1. DIRIGEANT: nom complet PDG/DG/Président + email professionnel + mobile si publics dans les sources.
   Site web officiel de l'entreprise.

2. FINANCIER — remplis tout ce que tu peux trouver:
   - ca_montant: valeur EXACTE du CA avec année (ex: "2.5 Md MAD (2023)", "450 M MAD (2022)").
     Si données Excel disponibles ci-dessus → utilise-les directement comme ca_montant.
     Si trouvé dans Tavily → utilise la valeur Tavily.
   - tranche_ca: DÉDUIS de ca_montant OU cherche: "< 10M MAD"|"10–50M MAD"|"50–100M MAD"|"100–500M MAD"|"> 500M MAD"
   - ca_export: montant export avec unité ET année si possible (ex: "120 M MAD (2022)"). Si seulement % → note le %.
   - part_export: pourcentage du CA exporté (ex: "30%").
   Si données Excel disponibles, confirme-les et enrichis avec l'année si possible.

3. EFFECTIF: tranche parmi "< 10"|"10–49"|"50–199"|"200–499"|"> 500". Si données Excel → utilise-les.

4. FILIÈRES — analyse rigoureuse basée sur les sources:
${hasSectors ? `   Pour chaque secteur coché (listés ci-dessus):
   - Confirmé par sources → filieres_confirmees
   - Non confirmé / incorrect → filieres_rejetees
   - Autres filières réelles non cochées → filieres_ajoutees` : `   - filieres_confirmees: filières prouvées
   - filieres_ajoutees: filières découvertes`}
   - filiere_principale: filière dominante (une seule), ou "Hors secteur FCP" si pharma/alimentaire/minier/pétrolier

5. TYPE — STRICTEMENT une de ces 8 valeurs EXACTES, rien d'autre, aucune variante:
   "Fabricant" | "Distributeur" | "Importateur" | "Importateur-Distributeur" | "Fabricant-Distributeur" | "Agent/Représentant" | "Négoce" | "Prestataire de services"
   ⚠️ INTERDIT: parenthèses, combinaisons libres, texte inventé. UNIQUEMENT ces 8 valeurs mot pour mot.

6. SECTEUR RÉEL: Si "Hors secteur FCP" → décris le vrai secteur (ex: "Industrie pharmaceutique", "Pétrole & hydrocarbures", "Industrie minière", "Agroalimentaire"). Sinon null.

7. CONFIANCE: 0.0–1.0 (pondère selon la qualité des données trouvées)

JSON uniquement (commence par {, sans backticks):
{
  "dirigeant": null,
  "email_dirigeant": null,
  "mobile_dirigeant": null,
  "site_web": null,
  "ca_montant": null,
  "tranche_ca": null,
  "ca_export": null,
  "part_export": null,
  "effectif_tranche": null,
  "filieres_confirmees": [],
  "filieres_rejetees": [],
  "filieres_ajoutees": [],
  "filiere_principale": null,
  "secteur_reel": null,
  "type_entreprise": null,
  "confiance": 0.0,
  "raison": "2-3 phrases avec preuves et sources utilisées",
  "sources": []
}`

  const client = new Anthropic({ apiKey: anthropicKey })
  const response = await client.messages.create({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 1000,
    messages: [{ role: 'user', content: prompt }],
  })

  const raw = response.content
    .filter((b) => b.type === 'text')
    .map((b) => (b as { text: string }).text)
    .join('')

  const j = raw.indexOf('{')
  const k = raw.lastIndexOf('}')
  if (j === -1 || k === -1) throw new Error(`JSON introuvable: ${raw.substring(0, 100)}`)
  const result = JSON.parse(raw.substring(j, k + 1))

  // ── Fallback: inject Excel financials if Claude missed them ──────────────
  if (!result.ca_montant && caFormatted)    result.ca_montant = caFormatted
  if (!result.tranche_ca && trancheFromExcel) result.tranche_ca = trancheFromExcel
  if (!result.ca_export  && caExpFormatted) result.ca_export  = caExpFormatted

  log(
    `  ✅ ${result.filiere_principale || '—'} | ${result.type_entreprise || '—'} | CA: ${result.ca_montant || '—'} | ${Math.round((result.confiance || 0) * 100)}%`
  )
  if (result.dirigeant) log(`  👤 ${result.dirigeant}`)
  if (result.site_web) log(`  🌐 ${result.site_web}`)

  // ── Reconciliation: every checked sector MUST get a verdict ───────────────
  // Any sector that was checked in the Excel but Claude left unaccounted
  // gets auto-added to filieres_rejetees (benefit of doubt: unconfirmed = rejected)
  if (company.sectorsChecked && company.sectorsChecked.length > 0) {
    const confirmed = new Set<string>(result.filieres_confirmees || [])
    const rejected  = new Set<string>(result.filieres_rejetees  || [])
    const unaccounted = company.sectorsChecked.filter(s => !confirmed.has(s) && !rejected.has(s))
    if (unaccounted.length > 0) {
      result.filieres_rejetees = [...(result.filieres_rejetees || []), ...unaccounted]
      log(`  ⚠ ${unaccounted.length} secteur(s) non tranchés → rejetés par défaut: ${unaccounted.join(', ')}`)
    }
  }

  const final = {
    code: company.code,
    name: company.name,
    city: company.city,
    sectorsChecked: company.sectorsChecked || [],   // keep original checked list
    ...result,
    processed_at: new Date().toISOString(),
  }

  // ── Optional Supabase save ───────────────────────────────────────────────
  if (supabaseUrl && supabaseKey && !supabaseUrl.includes('your_')) {
    try {
      const sb = createClient(supabaseUrl, supabaseKey)
      await sb
        .from('fcp_enrichments')
        .upsert({ ...final, updated_at: new Date().toISOString() }, { onConflict: 'code' })
      log(`  💾 Sauvegardé en DB`)
    } catch (e) {
      log(`  ⚠ DB: ${(e as Error).message}`)
    }
  }

  return { result: final, logs }
}

// ── Route handler ──────────────────────────────────────────────────────────
export async function POST(req: NextRequest) {
  const {
    companies,
    anthropicKey: anthropicKeyUI,
    tavilyKey: tavilyKeyUI,
    supabaseUrl: supabaseUrlUI,
    supabaseKey: supabaseKeyUI,
    concurrency = 5,
  } = await req.json()

  // Env vars take priority over UI-provided keys
  const anthropicKey  = process.env.ANTHROPIC_API_KEY  || anthropicKeyUI  || ''
  const tavilyKey     = process.env.TAVILY_API_KEY      || tavilyKeyUI     || ''
  const supabaseUrl   = process.env.SUPABASE_URL        || supabaseUrlUI   || ''
  const supabaseKey   = process.env.SUPABASE_KEY        || supabaseKeyUI   || ''

  if (!companies?.length || !anthropicKey || !tavilyKey) {
    return NextResponse.json({ error: 'Clés API manquantes (UI ou variables d\'environnement)' }, { status: 400 })
  }

  const results: Record<string, unknown>[] = []
  const allLogs: string[] = []

  for (let i = 0; i < companies.length; i += concurrency) {
    const batch = companies.slice(i, i + concurrency)
    const batchRes = await Promise.all(
      batch.map((c: CompanyInput) =>
        enrichOne(c, anthropicKey, tavilyKey, supabaseUrl, supabaseKey).catch((e) => ({
          result: {
            code: c.code,
            name: c.name,
            city: c.city,
            error: e.message,
            filieres_confirmees: [],
            filieres_rejetees: [],
            filieres_ajoutees: [],
            confiance: 0,
            processed_at: new Date().toISOString(),
          },
          logs: [`❌ ${c.name}: ${e.message}`],
        }))
      )
    )
    batchRes.forEach((r) => {
      results.push(r.result)
      allLogs.push(...r.logs)
    })
  }

  return NextResponse.json({ results, logs: allLogs })
}
