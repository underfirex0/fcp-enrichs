# FCP Intel

Agent IA d'enrichissement des données membres FCP — Claude Haiku + Tavily Search.

## Démarrage rapide (local)

```bash
npm install
cp .env.example .env.local   # édite avec tes clés
npm run dev
# → http://localhost:3000
```

## Déploiement Vercel + Supabase

### 1. Supabase
1. Crée un projet sur supabase.com
2. SQL Editor → exécute `supabase-schema.sql`
3. Récupère **Project URL** et **anon key**

### 2. Vercel — Variables d'environnement

| Variable | Description |
|---|---|
| `ANTHROPIC_API_KEY` | console.anthropic.com |
| `TAVILY_API_KEY` | tavily.com |
| `SUPABASE_URL` | https://xxx.supabase.co |
| `SUPABASE_KEY` | anon public key |

Push sur GitHub → import sur vercel.com/new → ajoute les 4 variables → Deploy.

## Comportement production
- Clés non affichées dans l'UI (🔒 configurées via env)
- Résultats chargés automatiquement depuis Supabase au démarrage
- Sauvegarde automatique après chaque enrichissement

## Format entrée
Upload direct du fichier `.xlsx` — lit Code, Nom, Ville, CA, CA Export, Effectif, Activité + secteurs cochés (X).

## Stack
Next.js 15 · Claude Haiku · Tavily Search · Supabase · Tailwind v4 · SheetJS
