import React, { createContext, useContext, useState, useEffect, ReactNode } from 'react';
import { SharePointService, ClientRecord } from '@/services/sharePointService';
import { authService } from '@/services/authService';
import { safeLocalStorage } from '@/lib/safe-storage';

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
  cachedClientRecords: ClientRecord[];
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
    const data = safeLocalStorage.getJSON<string[]>("sharepoint_cached_clients", []);
    if (Array.isArray(data) && data.length > 0) {
      console.log("🚀 SharePointContext initialized with cached clients:", data.length);
      return data;
    }
    if (!Array.isArray(data)) {
      console.warn("Cached clients data is not an array");
      return [];
    }
    return data;
  });
  const [cachedClientRecords, setCachedClientRecords] = useState<ClientRecord[]>(() => {
    const data = safeLocalStorage.getJSON<ClientRecord[]>("sharepoint_cached_client_records", []);
    if (Array.isArray(data) && data.length > 0) {
      console.log("🚀 SharePointContext initialized with cached client records:", data.length);
      return data;
    }
    if (!Array.isArray(data)) {
      console.warn("Cached client records data is not an array");
      return [];
    }
    return data;
  });
  const [cachedWorkOrders, setCachedWorkOrders] = useState<any[]>(() => {
    const data = safeLocalStorage.getJSON<any[]>("sharepoint_cached_workorders", []);
    if (Array.isArray(data) && data.length > 0) {
      console.log("🚀 SharePointContext initialized with cached work orders:", data.length);
      return data;
    }
    if (!Array.isArray(data)) {
      console.warn("Cached work orders data is not an array");
      return [];
    }
    return data;
  });
  const [isDataLoading, setIsDataLoading] = useState<boolean>(false);

  // Гипер-быстрое кеширование - сохраняем данные мгновенно с уведомлением всех компонентов
  const saveToCache = (key: string, data: any) => {
    const arrayData = Array.isArray(data) ? data : [];

    if (key === "sharepoint_cached_clients") {
      setCachedClients(arrayData);
    } else if (key === "sharepoint_cached_client_records") {
      setCachedClientRecords(arrayData as ClientRecord[]);
    } else if (key === "sharepoint_cached_workorders") {
      setCachedWorkOrders(arrayData);
    } else if (key === "sharepoint_cached_tubing" || key === "sharepoint_cached_sucker_rod" || key === "sharepoint_cached_coupling") {
      // Consumers read these through useSharePointInstantData, which taps localStorage directly.
    }

    try {
      safeLocalStorage.setJSON(key, data ?? []);
      safeLocalStorage.setItem(`${key}_timestamp`, new Date().toISOString());
      const serialized = JSON.stringify(data ?? []);
      console.log(`💾 Cached ${key}:`, Array.isArray(data) ? data.length : "data");
      safeLocalStorage.dispatchStorageEvent(key, serialized);
    } catch (error) {
      console.warn(`Failed to cache ${key}:`, error);
    }
  };

  // Проверяем свежесть кеша
  const isCacheFresh = (key: string, maxAgeMinutes = 30) => {
    const timestamp = safeLocalStorage.getItem(`${key}_timestamp`);
    if (!timestamp) return false;

    try {
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
        
        const storedConnection = safeLocalStorage.getItem("sharepoint_connected");
        
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
              safeLocalStorage.removeItem("sharepoint_connected");
              safeLocalStorage.removeItem("sharepoint_connection_time");
              setError('SharePoint authentication expired. Please reconnect.');
            }
            
          } else {
            console.log('No MSAL account found, clearing storage');
            safeLocalStorage.removeItem("sharepoint_connected");
            safeLocalStorage.removeItem("sharepoint_connection_time");
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
      const cachedClientsData = safeLocalStorage.getItem("sharepoint_cached_clients");
      const cachedClientRecordsData = safeLocalStorage.getItem("sharepoint_cached_client_records");
      const cachedWorkOrdersData = safeLocalStorage.getItem("sharepoint_cached_workorders");
      // tubing / sucker rod caches consumed via hooks
      
      if (cachedClientsData) {
        try {
          const clients = JSON.parse(cachedClientsData);
          if (Array.isArray(clients)) {
            setCachedClients(clients);
            if (clients.length > 0) {
              console.log('📦 Context loaded cached clients:', clients.length);
            }
          } else {
            console.warn('Cached clients data is not an array');
            setCachedClients([]);
          }
        } catch (parseError) {
          console.error('Error parsing cached clients data:', parseError);
          setCachedClients([]);
        }
      } else {
        setCachedClients([]);
      }

      if (cachedClientRecordsData) {
        try {
          const clientRecords = JSON.parse(cachedClientRecordsData);
          if (Array.isArray(clientRecords)) {
            setCachedClientRecords(clientRecords);
            if (clientRecords.length > 0) {
              console.log('📦 Context loaded cached client records:', clientRecords.length);
            }
          } else {
            console.warn('Cached client records data is not an array');
            setCachedClientRecords([]);
          }
        } catch (parseError) {
          console.error('Error parsing cached client records data:', parseError);
          setCachedClientRecords([]);
        }
      } else {
        setCachedClientRecords([]);
      }

      if (cachedWorkOrdersData) {
        try {
          const workOrders = JSON.parse(cachedWorkOrdersData);
          if (Array.isArray(workOrders)) {
            setCachedWorkOrders(workOrders);
            if (workOrders.length > 0) {
              console.log('📦 Context loaded cached work orders:', workOrders.length);
            }
          } else {
            console.warn('Cached work orders data is not an array');
            setCachedWorkOrders([]);
          }
        } catch (parseError) {
          console.error('Error parsing cached work orders data:', parseError);
          setCachedWorkOrders([]);
        }
      } else {
        setCachedWorkOrders([]);
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
      const lastRefresh = safeLocalStorage.getItem("sharepoint_last_refresh");
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

      try {
        const clientRecords = await service.getClientRecordsFromExcel();
        if (clientRecords && clientRecords.length > 0) {
          setCachedClientRecords(clientRecords);
          saveToCache('sharepoint_cached_client_records', clientRecords);
          console.log('✅ Cached detailed client records:', clientRecords.length);
        }
      } catch (error) {
        console.warn('Failed to load detailed client records:', error);
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

      // Загружаем tubing / sucker rod / coupling registry
      try {
        const tubingData = await service.getExcelData('tubing');
        if (tubingData && tubingData.length > 0) {
          saveToCache('sharepoint_cached_tubing', tubingData);
        }
      } catch (error) {
        console.warn('Failed to load tubing registry:', error);
      }

      try {
        const suckerRodData = await service.getExcelData('sucker_rod');
        if (suckerRodData && suckerRodData.length > 0) {
          saveToCache('sharepoint_cached_sucker_rod', suckerRodData);
        }
      } catch (error) {
        console.warn('Failed to load sucker rod registry:', error);
      }

      try {
        const couplingData = await service.getExcelData('coupling');
        if (couplingData && couplingData.length > 0) {
          saveToCache('sharepoint_cached_coupling', couplingData);
        }
      } catch (error) {
        console.warn('Failed to load coupling registry:', error);
      }

      safeLocalStorage.setItem("sharepoint_last_refresh", now.toISOString());
      
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
      safeLocalStorage.setItem("sharepoint_connected", "true");
      safeLocalStorage.setItem("sharepoint_connection_time", new Date().toISOString());
      
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
    setCachedClientRecords([]);
    setCachedWorkOrders([]);

    // Очищаем сохраненное состояние SharePoint (но НЕ MSAL токены)
    [
      "sharepoint_connected",
      "sharepoint_connection_time",
      "sharepoint_cached_clients",
      "sharepoint_cached_client_records",
      "sharepoint_cached_workorders",
      "sharepoint_cached_tubing",
      "sharepoint_cached_sucker_rod",
      "sharepoint_clients_timestamp",
      "sharepoint_cached_client_records_timestamp",
      "sharepoint_workorders_timestamp",
      "sharepoint_cached_tubing_timestamp",
      "sharepoint_last_refresh",
    ].forEach(key => safeLocalStorage.removeItem(key));

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
    cachedClientRecords,
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
