// Data Types for Form-to-Table App

export interface WorkOrder {
  id: string;
  title: string;
  wo_num: number;
  client: string;
  pipe_type: string;
  diameter: string;
  coupling_replace: boolean;
  transport: string;
  status: 'Open' | 'In Progress' | 'Completed';
  created_at: string;
  added_by: string;
}

export interface BatchRegistry {
  id: string;
  title: string;
  wo_id: string;
  wo_num: number;
  client: string;
  batch_no: number;
  qty: number;
  status: 'RECEIVED, AWAITING INSPECTION' | 'INSPECTION DONE' | 'AWAITING AVR' | 'AVR DONE';
  rack: string;
  registered_at: string;
  note?: string;
  order_key: string;
  added_by: string;
}

export interface InspectionData {
  id: string;
  wo_id: string;
  batch_id: string;
  class1: number;
  class2: number;
  repair: number;
  scrap_total: number;
  rattling_qty: number;
  external_qty: number;
  hydro_qty: number;
  mpi_qty: number;
  drift_qty: number;
  emi_qty: number;
  inspector_name: string;
  completed_at: string;
}

export interface LoadOutLog {
  id: string;
  wo_id: string;
  batch_id: string;
  load_list_no: string;
  load_out_date: string;
  comment?: string;
}

export interface AVRLog {
  id: string;
  wo_id: string;
  batch_id: string;
  avr_no: string;
  avr_date: string;
  comment?: string;
  attachment_url?: string;
}

export interface BoardItem {
  id: string;
  rack: string;
  client: string;
  batch: string;
  wo: string;
  qty: number;
  status: string;
  order_key: string;
  registered_at: string;
}

export interface DashboardStats {
  totalRecords2025: number;
  awaitingInspection: number;
  inspectionDone: number;
  awaitingAVR: number;
  completed: number;
}

// Form Schemas
export interface CreateWOFormData {
  wo_num: number;
  client: string;
  pipe_type: string;
  diameter: string;
  coupling_replace: boolean;
  transport: string;
}

export interface BatchRegistryFormData {
  wo_id: string;
  wo_num: number;
  client: string;
  qty: number;
  rack: string;
  note?: string;
}

export interface InspectionFormData {
  batch_id: string;
  class1: number;
  class2: number;
  repair: number;
  scrap_total: number;
  rattling_qty: number;
  external_qty: number;
  hydro_qty: number;
  mpi_qty: number;
  drift_qty: number;
  emi_qty: number;
}

export interface LoadOutFormData {
  load_list_no: string;
  load_out_date: string;
  comment?: string;
}

export interface AVRFormData {
  avr_no: string;
  avr_date: string;
  comment?: string;
  attachment?: File;
}