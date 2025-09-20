import React, { createContext, useContext, useState, useEffect, ReactNode } from 'react';
import { SharePointService } from '@/services/sharePointService';
import { authService } from '@/services/authService';

interface SharePointContextType {
  isConnected: boolean;
  isConnecting: boolean;
  sharePointService: SharePointService | null;
  connect: () => Promise<boolean>;
  connectToSharePoint: () => Promise<boolean>;
  disconnect: () => void;
  error: string | null;
  // Кешированные данные
  cachedClients: string[];
  cachedWorkOrders: any[];
  isDataLoading: boolean;
  refreshData: () => Promise<void>;
  refreshDataInBackground: (service: SharePointService) => Promise<void>;
  resetExcelSession: () => void;
}

const SharePointContext = createContext<SharePointContextType | undefined>(undefined);

export const useSharePoint = () => {
  const context = useContext(SharePointContext);
  if (context === undefined) {
    throw new Error('useSharePoint must be used within a SharePointProvider');
  }
  return context;
};

interface SharePointProviderProps {
  children: ReactNode;
}

export const SharePointProvider: React.FC<SharePointProviderProps> = ({ children }) => {
  const [isConnected, setIsConnected] = useState<boolean>(false);
  const [isConnecting, setIsConnecting] = useState<boolean>(false);
  const [sharePointService, setSharePointService] = useState<SharePointService | null>(null);
  const [error, setError] = useState<string | null>(null);
  
  // Кешированные данные для быстрого доступа - инициализируем сразу из localStorage
  const [cachedClients, setCachedClients] = useState<string[]>(() => {
    try {
      const cached = localStorage.getItem('sharepoint_cached_clients');
      const data = cached ? JSON.parse(cached) : [];
      console.log('🚀 SharePointContext initialized with cached clients:', data.length);
      return data;
    } catch { 
      console.warn('Failed to load cached clients');
      return []; 
    }
  });
  const [cachedWorkOrders, setCachedWorkOrders] = useState<any[]>(() => {
    try {
      const cached = localStorage.getItem('sharepoint_cached_workorders');
      const data = cached ? JSON.parse(cached) : [];
      console.log('🚀 SharePointContext initialized with cached work orders:', data.length);
      return data;
    } catch { 
      console.warn('Failed to load cached work orders');
      return []; 
    }
  });
  const [isDataLoading, setIsDataLoading] = useState<boolean>(false);

  // Гипер-быстрое кеширование - сохраняем данные мгновенно с уведомлением всех компонентов
  const saveToCache = (key: string, data: any) => {
    try {
      localStorage.setItem(key, JSON.stringify(data));
      localStorage.setItem(`${key}_timestamp`, new Date().toISOString());
      console.log(`💾 Cached ${key}:`, data.length || 'data');
      
      // Обновляем локальное состояние для мгновенного отображения
      if (key === 'sharepoint_cached_clients') {
        setCachedClients(data);
      } else if (key === 'sharepoint_cached_workorders') {
        setCachedWorkOrders(data);
      }
      
      // Уведомляем все компоненты о изменении данных
      window.dispatchEvent(new StorageEvent('storage', {
        key: key,
        newValue: JSON.stringify(data),
        storageArea: localStorage
      }));
    } catch (error) {
      console.warn(`Failed to cache ${key}:`, error);
    }
  };

  // Проверяем свежесть кеша
  const isCacheFresh = (key: string, maxAgeMinutes = 30) => {
    try {
      const timestamp = localStorage.getItem(`${key}_timestamp`);
      if (!timestamp) return false;
      
      const cacheTime = new Date(timestamp);
      const now = new Date();
      const ageMinutes = (now.getTime() - cacheTime.getTime()) / (1000 * 60);
      
      return ageMinutes < maxAgeMinutes;
    } catch {
      return false;
    }
  };

  // Проверяем сохраненное состояние аутентификации при загрузке
  useEffect(() => {
    const checkStoredAuth = async () => {
      try {
        // ВСЕГДА загружаем кешированные данные первым делом
        loadCachedData();
        
        const storedConnection = localStorage.getItem('sharepoint_connected');
        
        if (storedConnection === 'true') {
          console.log('Found stored SharePoint connection, attempting to restore...');
          
          // Инициализируем auth service
          await authService.initialize();
          
          // Проверяем есть ли активные аккаунты MSAL
          if (authService.isAuthenticated()) {
            console.log('MSAL account found, creating SharePoint service...');
            
            try {
              // Создаем SharePoint service с auth provider
              const authProvider = authService.getAuthenticationProvider();
              const service = new SharePointService(authProvider);
              
              // Проверяем что токен работает - делаем легкую проверку
              console.log('🔍 Testing SharePoint connection...');
              await service.testConnection();
              
              setSharePointService(service);
              setIsConnected(true);
              console.log('✅ SharePoint authentication restored and verified');
              
              // Запускаем фоновое обновление данных БЕЗ блокировки UI
              setTimeout(() => {
                refreshDataInBackground(service);
              }, 100); // Задержка 100мс для мгновенного обновления
            } catch (connectionError) {
              console.error('❌ SharePoint connection test failed:', connectionError);
              console.log('⚠️ Clearing invalid SharePoint session');
              localStorage.removeItem('sharepoint_connected');
              localStorage.removeItem('sharepoint_connection_time');
              setError('SharePoint authentication expired. Please reconnect.');
            }
            
          } else {
            console.log('No MSAL account found, clearing storage');
            localStorage.removeItem('sharepoint_connected');
            localStorage.removeItem('sharepoint_connection_time');
          }
        }
      } catch (error) {
        console.error('Error checking stored auth:', error);
        // Все равно загружаем кеш даже при ошибках
        loadCachedData();
      }
    };

    checkStoredAuth();
  }, []);

  // Загрузка кешированных данных из localStorage
  const loadCachedData = () => {
    try {
      const cachedClientsData = localStorage.getItem('sharepoint_cached_clients');
      const cachedWorkOrdersData = localStorage.getItem('sharepoint_cached_workorders');
      
      if (cachedClientsData) {
        const clients = JSON.parse(cachedClientsData);
        if (clients.length > 0) {
          setCachedClients(clients);
          console.log('📦 Context loaded cached clients:', clients.length);
        }
      }
      
      if (cachedWorkOrdersData) {
        const workOrders = JSON.parse(cachedWorkOrdersData);
        if (workOrders.length > 0) {
          setCachedWorkOrders(workOrders);
          console.log('📦 Context loaded cached work orders:', workOrders.length);
        }
      }
    } catch (error) {
      console.error('Error loading cached data:', error);
    }
  };

  // Фоновое обновление данных
  const refreshDataInBackground = async (service: SharePointService) => {
    try {
      setIsDataLoading(true);
      console.log('🔄 Starting background data refresh...');
      
      // Проверяем время последнего обновления
      const lastRefresh = localStorage.getItem('sharepoint_last_refresh');
      const now = new Date();
      const lastRefreshTime = lastRefresh ? new Date(lastRefresh) : null;
      
      // Обновляем данные только если прошло больше 5 секунд (для real-time)
      // НО если время последнего обновления было очищено - всегда обновляем
      if (lastRefreshTime && (now.getTime() - lastRefreshTime.getTime()) < 5 * 1000) {
        console.log('📦 Data is fresh, skipping refresh (use Update Data to force refresh)');
        setIsDataLoading(false);
        return;
      }
      
      // Загружаем клиентов
      console.log('🔄 Making SharePoint API call to get clients...');
      const clients = await service.getClients();
      console.log('📊 SharePoint API returned clients:', clients?.length || 0, clients);
      if (clients && clients.length > 0) {
        setCachedClients(clients);
        saveToCache('sharepoint_cached_clients', clients);
        console.log('✅ Successfully cached', clients.length, 'clients');
      } else {
        console.warn('⚠️ No clients received from SharePoint API');
      }
      
      // Загружаем полные данные work orders из Excel листа 'wo'
      try {
        const workOrdersData = await service.getExcelData('wo');
        if (workOrdersData && workOrdersData.length > 0) {
          setCachedWorkOrders(workOrdersData);
          saveToCache('sharepoint_cached_workorders', workOrdersData);
          console.log('✅ Successfully cached full work orders data:', workOrdersData.length, 'rows');
        }
      } catch (error) {
        console.warn('Failed to load work orders data:', error);
      }

      // Загружаем tubing registry для SharePoint Viewer
      try {
        const tubingData = await service.getExcelData('tubing');
        if (tubingData && tubingData.length > 0) {
          saveToCache('sharepoint_cached_tubing', tubingData);
        }
      } catch (error) {
        console.warn('Failed to load tubing registry:', error);
      }
      
      // Сохраняем время последнего обновления
      localStorage.setItem('sharepoint_last_refresh', now.toISOString());
      
    } catch (error) {
      console.error('Background data refresh failed:', error);
    } finally {
      setIsDataLoading(false);
    }
  };

  // Принудительное обновление данных
  const refreshData = async () => {
    if (!sharePointService) return;
    await refreshDataInBackground(sharePointService);
  };

  // Ручной сброс Excel session (Workbook Session ID)
  const resetExcelSession = () => {
    try {
      if (sharePointService) {
        console.log('🔁 Resetting Excel workbook session by user request...');
        sharePointService.resetExcelSession();
      } else {
        console.warn('Cannot reset Excel session: SharePoint service is not connected');
      }
    } catch (error) {
      console.error('Failed to reset Excel session:', error);
    }
  };

  const connect = async (): Promise<boolean> => {
    if (isConnecting) return false;
    
    setIsConnecting(true);
    setError(null);

    try {
      await authService.initialize();
      
      // Сначала пробуем тихий вход
      let success = await authService.login();
      
      // Если тихий вход не сработал, используем popup
      if (!success) {
        success = await authService.forcePopupLogin();
      }
      
      if (!success) throw new Error('Authentication failed');
      
      const authProvider = authService.getAuthenticationProvider();
      const service = new SharePointService(authProvider);
      
      // НЕ тестируем подключение - доверяем MSAL токенам
      setSharePointService(service);
      setIsConnected(true);
      
      // Сохраняем состояние в localStorage
      localStorage.setItem('sharepoint_connected', 'true');
      localStorage.setItem('sharepoint_connection_time', new Date().toISOString());
      
      console.log('✅ SharePoint connected and saved to storage');
      
      // Загружаем кешированные данные
      loadCachedData();
      
      // Запускаем фоновое обновление данных с задержкой
      setTimeout(() => {
        refreshDataInBackground(service);
      }, 50); // 50мс для мгновенного обновления
      
      return true;
    } catch (error: any) {
      console.error('SharePoint connection failed:', error);
      setError(error.message || 'Ошибка подключения к SharePoint');
      return false;
    } finally {
      setIsConnecting(false);
    }
  };

  const disconnect = () => {
    setSharePointService(null);
    setIsConnected(false);
    setError(null);
    setCachedClients([]);
    setCachedWorkOrders([]);
    
    // Очищаем сохраненное состояние SharePoint (но НЕ MSAL токены)
    localStorage.removeItem('sharepoint_connected');
    localStorage.removeItem('sharepoint_connection_time');
    localStorage.removeItem('sharepoint_cached_clients');
    localStorage.removeItem('sharepoint_cached_workorders');
    localStorage.removeItem('sharepoint_clients_timestamp');
    localStorage.removeItem('sharepoint_workorders_timestamp');
    localStorage.removeItem('sharepoint_last_refresh');
    
    console.log('SharePoint disconnected and cache cleared (MSAL tokens preserved)');
  };

  const value = {
    isConnected,
    isConnecting,
    sharePointService,
    connect,
    connectToSharePoint: connect,
    disconnect,
    error,
    cachedClients,
    cachedWorkOrders,
    isDataLoading,
    refreshData,
    refreshDataInBackground,
    resetExcelSession
  };

  return (
    <SharePointContext.Provider value={value}>
      {children}
    </SharePointContext.Provider>
  );
};
