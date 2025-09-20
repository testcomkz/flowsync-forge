import { supabase } from '@/lib/supabase';
import type { WorkOrder, TubingRecord, Client } from '@/lib/supabase';

export class SupabaseService {
  // User authentication
  async authenticateUser(email: string, password: string): Promise<{ success: boolean; user?: any; error?: string }> {
    try {
      console.log('Authenticating user:', email);
      
      const { data, error } = await supabase
        .from('app_users')
        .select('*')
        .eq('email', email)
        .eq('password', password)
        .eq('is_active', true)
        .single();

      console.log('Auth result:', { data, error });

      if (error || !data) {
        console.log('Auth failed:', error?.message);
        return { success: false, error: 'Неверный email или пароль' };
      }

      console.log('Auth success:', data);
      return { success: true, user: data };
    } catch (error) {
      console.log('Auth exception:', error);
      return { success: false, error: 'Ошибка подключения к базе данных' };
    }
  }

  // Получить список клиентов
  async getClients(): Promise<string[]> {
    try {
      const { data, error } = await supabase
        .from('clients')
        .select('name')
        .order('name');

      if (error) throw error;
      return data?.map(client => client.name) || [];
    } catch (error) {
      console.error('Error fetching clients:', error);
      return ['Dunga', 'KenSary', 'Tasbulat']; // Fallback data
    }
  }

  // Получить Work Orders для клиента
  async getWorkOrdersByClient(client: string): Promise<string[]> {
    try {
      const { data, error } = await supabase
        .from('work_orders')
        .select('wo_no')
        .eq('client', client)
        .order('wo_no');

      if (error) throw error;
      return data?.map(wo => wo.wo_no) || [];
    } catch (error) {
      console.error('Error fetching work orders:', error);
      return [];
    }
  }

  // Получить диаметры для Work Order
  async getDiametersByWorkOrder(client: string, woNo: string): Promise<string[]> {
    try {
      const { data, error } = await supabase
        .from('work_orders')
        .select('diameter')
        .eq('client', client)
        .eq('wo_no', woNo);

      if (error) throw error;
      return data?.map(wo => wo.diameter) || [];
    } catch (error) {
      console.error('Error fetching diameters:', error);
      return [];
    }
  }

  // Получить информацию о батчах
  async getBatchInfo(client: string, woNo: string): Promise<{ lastBatch: number; lastPipeTo: number }> {
    try {
      const { data, error } = await supabase
        .from('tubing_registry')
        .select('batch, pipe_to')
        .eq('client', client)
        .eq('wo_no', woNo)
        .order('sort_order', { ascending: false })
        .limit(1);

      if (error) throw error;

      if (data && data.length > 0) {
        const lastItem = data[0];
        return {
          lastBatch: parseInt(lastItem.batch.replace('Batch # ', '')),
          lastPipeTo: lastItem.pipe_to
        };
      }

      return { lastBatch: 0, lastPipeTo: 0 };
    } catch (error) {
      console.error('Error fetching batch info:', error);
      return { lastBatch: 0, lastPipeTo: 0 };
    }
  }

  // Создать новый Work Order
  async createWorkOrder(data: Partial<WorkOrder>): Promise<boolean> {
    try {
      const { error } = await supabase
        .from('work_orders')
        .insert([{
          wo_no: data.wo_no,
          client: data.client,
          type: data.type,
          diameter: data.diameter,
          coupling_replace: data.coupling_replace,
          wo_date: data.wo_date,
          transport: data.transport,
          key_col: data.key_col,
          payer: data.payer,
          planned_qty: data.planned_qty
        }]);

      if (error) throw error;
      return true;
    } catch (error) {
      console.error('Error creating work order:', error);
      return false;
    }
  }

  // Создать запись в Tubing Registry с правильным порядком вставки
  async createTubingRecord(data: Partial<TubingRecord>): Promise<boolean> {
    try {
      // Найти позицию для вставки
      const insertPosition = await this.findInsertPosition(data.client!, data.wo_no!);
      
      const { error } = await supabase
        .from('tubing_registry')
        .insert([{
          client: data.client,
          wo_no: data.wo_no,
          batch: data.batch,
          diameter: data.diameter,
          qty: data.qty,
          pipe_from: data.pipe_from,
          pipe_to: data.pipe_to,
          class_1: data.class_1,
          class_2: data.class_2,
          class_3: data.class_3,
          repair: data.repair,
          scrap: data.scrap,
          start_date: data.start_date,
          end_date: data.end_date,
          rattling_qty: data.rattling_qty || 0,
          external_qty: data.external_qty || 0,
          hydro_qty: data.hydro_qty || 0,
          mpi_qty: data.mpi_qty || 0,
          drift_qty: data.drift_qty || 0,
          emi_qty: data.emi_qty || 0,
          marking_qty: data.marking_qty || 0,
          act_no_oper: data.act_no_oper,
          act_date: data.act_date,
          sort_order: insertPosition
        }]);

      if (error) throw error;
      return true;
    } catch (error) {
      console.error('Error creating tubing record:', error);
      return false;
    }
  }

  // Найти правильную позицию для вставки записи
  private async findInsertPosition(client: string, woNo: string): Promise<number> {
    try {
      // Получить последнюю запись этого клиента и WO
      const { data: clientRecords } = await supabase
        .from('tubing_registry')
        .select('sort_order')
        .eq('client', client)
        .eq('wo_no', woNo)
        .order('sort_order', { ascending: false })
        .limit(1);

      if (clientRecords && clientRecords.length > 0) {
        return clientRecords[0].sort_order + 1;
      }

      // Если нет записей для этого клиента и WO, найти последнюю запись клиента
      const { data: lastClientRecord } = await supabase
        .from('tubing_registry')
        .select('sort_order')
        .eq('client', client)
        .order('sort_order', { ascending: false })
        .limit(1);

      if (lastClientRecord && lastClientRecord.length > 0) {
        return lastClientRecord[0].sort_order + 1;
      }

      // Если это первая запись клиента, найти общую последнюю позицию
      const { data: allRecords } = await supabase
        .from('tubing_registry')
        .select('sort_order')
        .order('sort_order', { ascending: false })
        .limit(1);

      return allRecords && allRecords.length > 0 ? allRecords[0].sort_order + 100 : 1000;
    } catch (error) {
      console.error('Error finding insert position:', error);
      return Date.now(); // Fallback to timestamp
    }
  }

  // Получить все записи для редактирования
  async getTubingRecords(filters?: { client?: string; wo_no?: string; batch?: string }): Promise<TubingRecord[]> {
    try {
      let query = supabase
        .from('tubing_registry')
        .select('*')
        .order('sort_order');

      if (filters?.client) {
        query = query.ilike('client', `%${filters.client}%`);
      }
      if (filters?.wo_no) {
        query = query.ilike('wo_no', `%${filters.wo_no}%`);
      }
      if (filters?.batch) {
        query = query.ilike('batch', `%${filters.batch}%`);
      }

      const { data, error } = await query;
      if (error) throw error;
      return data || [];
    } catch (error) {
      console.error('Error fetching tubing records:', error);
      return [];
    }
  }

  // Обновить запись
  async updateTubingRecord(id: number, data: Partial<TubingRecord>): Promise<boolean> {
    try {
      const { error } = await supabase
        .from('tubing_registry')
        .update(data)
        .eq('id', id);

      if (error) throw error;
      return true;
    } catch (error) {
      console.error('Error updating tubing record:', error);
      return false;
    }
  }
}
