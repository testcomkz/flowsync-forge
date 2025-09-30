import { supabase } from '@/lib/supabase';
import { safeLocalStorage } from '@/lib/safe-storage';
import type { RealtimeChannel } from '@supabase/supabase-js';

class SupabaseRealtimeService {
  private channel: RealtimeChannel | null = null;
  private listeners: Map<string, Set<Function>> = new Map();

  // Подключение к Realtime каналу (используем broadcast вместо postgres_changes)
  connect(): void {
    if (this.channel) {
      console.log('✅ Already subscribed to Realtime');
      return;
    }

    console.log('🔌 Connecting to Supabase Realtime...');

    this.channel = supabase
      .channel('sharepoint-updates')
      .on('broadcast', { event: 'excel-updated' }, (payload) => {
        console.log('📊 Realtime update received:', payload);
        this.handleUpdate(payload.payload);
      })
      .subscribe((status) => {
        if (status === 'SUBSCRIBED') {
          console.log('✅ Subscribed to Supabase Realtime');
        } else if (status === 'CHANNEL_ERROR') {
          console.error('❌ Realtime subscription error');
        }
      });
  }

  // Обработка обновления
  private async handleUpdate(update: any): Promise<void> {
    console.log('🔄 Processing SharePoint update...');

    try {
      // Получаем обновленные данные из SharePoint через существующий сервис
      // Триггерим обновление через storage event
      safeLocalStorage.setItem('sharepoint_force_refresh', Date.now().toString());
      safeLocalStorage.dispatchStorageEvent('sharepoint_force_refresh', Date.now().toString());

      // Уведомляем подписчиков
      this.notifyListeners('data-updated', update);

      console.log('✅ Update processed');
    } catch (error) {
      console.error('❌ Error processing update:', error);
    }
  }

  // Подписка на события
  on(event: string, callback: Function): void {
    if (!this.listeners.has(event)) {
      this.listeners.set(event, new Set());
    }
    this.listeners.get(event)!.add(callback);
  }

  // Отписка от событий
  off(event: string, callback: Function): void {
    const callbacks = this.listeners.get(event);
    if (callbacks) {
      callbacks.delete(callback);
    }
  }

  // Уведомление подписчиков
  private notifyListeners(event: string, data: any): void {
    const callbacks = this.listeners.get(event);
    if (callbacks) {
      callbacks.forEach(callback => callback(data));
    }
  }

  // Отключение
  disconnect(): void {
    if (this.channel) {
      supabase.removeChannel(this.channel);
      this.channel = null;
      console.log('👋 Disconnected from Realtime');
    }
  }

  // Проверка подключения
  isConnected(): boolean {
    return this.channel !== null;
  }
}

export const supabaseRealtimeService = new SupabaseRealtimeService();
