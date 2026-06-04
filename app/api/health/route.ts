import { NextResponse } from 'next/server'

export async function GET() {
  return NextResponse.json({
    anthropic: !!process.env.ANTHROPIC_API_KEY,
    tavily:    !!process.env.TAVILY_API_KEY,
    supabase:  !!(process.env.SUPABASE_URL && process.env.SUPABASE_KEY),
  })
}
