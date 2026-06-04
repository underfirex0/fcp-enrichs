import { NextRequest, NextResponse } from 'next/server'
import { createClient } from '@supabase/supabase-js'

export async function POST(req: NextRequest) {
  const body = await req.json().catch(() => ({}))
  // Env vars take priority
  const supabaseUrl = process.env.SUPABASE_URL || body.supabaseUrl
  const supabaseKey = process.env.SUPABASE_KEY || body.supabaseKey

  if (!supabaseUrl || !supabaseKey) {
    return NextResponse.json({ error: 'Supabase credentials manquantes' }, { status: 400 })
  }
  try {
    const sb = createClient(supabaseUrl, supabaseKey)
    const { data, error } = await sb
      .from('fcp_enrichments')
      .select('*')
      .order('updated_at', { ascending: false })
    if (error) throw error
    return NextResponse.json({ results: data || [] })
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 500 })
  }
}
