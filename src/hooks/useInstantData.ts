import { useState, useEffect } from 'react';
import { safeLocalStorage } from '@/lib/safe-storage';
import type { ClientRecord } from '@/services/sharePointService';

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
  const [clientRecords, setClientRecords] = useInstantData<ClientRecord[]>('sharepoint_cached_client_records', []);
  const [workOrders, setWorkOrders] = useInstantData<any[]>('sharepoint_cached_workorders', []);
  const [tubingData, setTubingData] = useInstantData<any[]>('sharepoint_cached_tubing', []);
  const [suckerRodData, setSuckerRodData] = useInstantData<any[]>('sharepoint_cached_sucker_rod', []);
  const [couplingData, setCouplingData] = useInstantData<any[]>('sharepoint_cached_coupling', []);
  
  // Force re-sync on component mount to ensure fresh data
  useEffect(() => {
    const syncData = () => {
      try {
        const cachedClients = safeLocalStorage.getItem("sharepoint_cached_clients");
        const cachedClientRecords = safeLocalStorage.getItem("sharepoint_cached_client_records");
        const cachedWorkOrders = safeLocalStorage.getItem("sharepoint_cached_workorders");
        const cachedTubing = safeLocalStorage.getItem("sharepoint_cached_tubing");
        const cachedSuckerRod = safeLocalStorage.getItem("sharepoint_cached_sucker_rod");
        const cachedCoupling = safeLocalStorage.getItem("sharepoint_cached_coupling");
        
        if (cachedClients) {
          const clientsData = JSON.parse(cachedClients);
          if (clientsData.length > 0) setClients(clientsData);
        }
        
        if (cachedClientRecords) {
          try {
            const records = JSON.parse(cachedClientRecords);
            if (Array.isArray(records) && records.length > 0) {
              setClientRecords(records);
            }
          } catch (error) {
            console.warn('Failed to sync client records from cache:', error);
          }
        }

        if (cachedWorkOrders) {
          const workOrdersData = JSON.parse(cachedWorkOrders);
          if (workOrdersData.length > 0) setWorkOrders(workOrdersData);
        }
        
        if (cachedTubing) {
          const tubingDataParsed = JSON.parse(cachedTubing);
          if (tubingDataParsed.length > 0) setTubingData(tubingDataParsed);
        }

        if (cachedSuckerRod) {
          const suckerRodParsed = JSON.parse(cachedSuckerRod);
          if (suckerRodParsed.length > 0) setSuckerRodData(suckerRodParsed);
        }

        if (cachedCoupling) {
          const couplingParsed = JSON.parse(cachedCoupling);
          if (couplingParsed.length > 0) setCouplingData(couplingParsed);
        }
      } catch (error) {
        console.warn('Failed to sync SharePoint data:', error);
      }
    };

    syncData();
  }, [setClients, setClientRecords, setWorkOrders, setTubingData, setSuckerRodData, setCouplingData]);

  return { clients, clientRecords, workOrders, tubingData, suckerRodData, couplingData };
}
