import { NextRequest, NextResponse } from 'next/server'
import { createClient } from '@supabase/supabase-js'

export async function POST(req: NextRequest) {
  const { results, supabaseUrl: urlUI, supabaseKey: keyUI } = await req.json()
  const url = process.env.SUPABASE_URL || urlUI
  const key = process.env.SUPABASE_KEY || keyUI
  if (!url || !key) return NextResponse.json({ error: 'Supabase non configuré' }, { status: 400 })
  if (!results?.length) return NextResponse.json({ error: 'Aucun résultat' }, { status: 400 })
  try {
    const sb = createClient(url, key)
    const { error } = await sb.from('fcp_enrichments').upsert(
      results.map((r: any) => ({ ...r, updated_at: new Date().toISOString() })),
      { onConflict: 'code' }
    )
    if (error) throw error
    return NextResponse.json({ saved: results.length })
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 500 })
  }
}
