import { useState, useEffect } from 'react';
import { safeLocalStorage } from '@/lib/safe-storage';

// Ultra-fast data hook for millisecond loading with automatic sync
export const useInstantData = <T>(key: string, defaultValue: T) => {
  const [data, setData] = useState<T>(() => {
    try {
      const cached = safeLocalStorage.getItem(key);
      return cached ? JSON.parse(cached) : defaultValue;
    } catch {
      return defaultValue;
    }
  });

  // Listen for storage changes to sync data across components
  useEffect(() => {
    const handleStorageChange = (e: StorageEvent) => {
      if (e.key === key && e.newValue) {
        try {
          const newData = JSON.parse(e.newValue);
          setData(newData);
        } catch (error) {
          console.warn(`Failed to parse storage change for ${key}:`, error);
        }
      }
    };

    window.addEventListener('storage', handleStorageChange);
    return () => window.removeEventListener('storage', handleStorageChange);
  }, [key]);

  const updateData = (newData: T) => {
    setData(newData);
    try {
      safeLocalStorage.setItem(key, JSON.stringify(newData));
    } catch (error) {
      console.warn(`Failed to save ${key} to localStorage:`, error);
    }
  };

  return [data, updateData] as const;
};

// Pre-load all SharePoint data instantly with persistent sync
export const useSharePointInstantData = () => {
  const [clients, setClients] = useInstantData<string[]>('sharepoint_cached_clients', []);
  const [workOrders, setWorkOrders] = useInstantData<any[]>('sharepoint_cached_workorders', []);
  const [tubingData, setTubingData] = useInstantData<any[]>('sharepoint_cached_tubing', []);
  
  // Force re-sync on component mount to ensure fresh data
  useEffect(() => {
    const syncData = () => {
      try {
        const cachedClients = safeLocalStorage.getItem("sharepoint_cached_clients");
        const cachedWorkOrders = safeLocalStorage.getItem("sharepoint_cached_workorders");
        const cachedTubing = safeLocalStorage.getItem("sharepoint_cached_tubing");
        
        if (cachedClients) {
          const clientsData = JSON.parse(cachedClients);
          if (clientsData.length > 0) setClients(clientsData);
        }
        
        if (cachedWorkOrders) {
          const workOrdersData = JSON.parse(cachedWorkOrders);
          if (workOrdersData.length > 0) setWorkOrders(workOrdersData);
        }
        
        if (cachedTubing) {
          const tubingDataParsed = JSON.parse(cachedTubing);
          if (tubingDataParsed.length > 0) setTubingData(tubingDataParsed);
        }
      } catch (error) {
        console.warn('Failed to sync SharePoint data:', error);
      }
    };

    syncData();
  }, [setClients, setWorkOrders, setTubingData]);
  
  return { clients, workOrders, tubingData };
};
