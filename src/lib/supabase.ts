import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://hkndhnctlctvybfftyqu.supabase.co'
const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImhrbmRobmN0bGN0dnliZmZ0eXF1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTUyNDI4OTEsImV4cCI6MjA3MDgxODg5MX0.G9DMDu3jiSiCR4eMMgJdpim778OQ0vNgiLlo_ViA2CA'

console.log('Supabase URL:', supabaseUrl)
console.log('Supabase Key:', supabaseAnonKey.substring(0, 20) + '...')

export const supabase = createClient(supabaseUrl, supabaseAnonKey)

// Database types
export interface WorkOrder {
  id: number
  wo_no: string
  client: string
  type: string
  diameter: string
  coupling_replace?: string
  wo_date: string
  transport?: string
  key_col?: string
  payer?: string
  planned_qty: number
  created_at: string
}

export interface TubingRecord {
  id: number
  client: string
  wo_no: string
  batch: string
  diameter: string
  qty: number
  pipe_from: number
  pipe_to: number
  class_1?: string
  class_2?: string
  class_3?: string
  repair?: string
  scrap?: string
  start_date?: string
  end_date?: string
  rattling_qty?: number
  external_qty?: number
  hydro_qty?: number
  mpi_qty?: number
  drift_qty?: number
  emi_qty?: number
  marking_qty?: number
  act_no_oper?: string
  act_date?: string
  sort_order: number
  created_at: string
}

export interface Client {
  id: number
  name: string
  created_at: string
}
