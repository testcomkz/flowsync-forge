import React, { useEffect, useMemo, useState } from 'react';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Badge } from "@/components/ui/badge";
import { Dialog, DialogContent, DialogDescription, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { ToggleGroup, ToggleGroupItem } from "@/components/ui/toggle-group";
import { Input } from "@/components/ui/input";
import { Loader2, FileSpreadsheet, Users, Briefcase, Database, ArrowLeft, Info, Filter } from "lucide-react";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useToast } from "@/hooks/use-toast";
import { useNavigate } from "react-router-dom";
import { safeLocalStorage } from '@/lib/safe-storage';

interface ExcelData {
  clients: string[];
  workOrders: unknown[];
  tubingRegistry: unknown[];
}

type ParsedRow = {
  values: unknown[];
  data: Record<string, unknown>;
};

interface ParsedSheetData {
  headers: string[];
  canonicalHeaders: string[];
  rows: ParsedRow[];
}

type StatusKey = 'arrived' | 'inspection_done' | 'completed' | 'other';

interface WorkOrderEntry {
  id: string;
  client: string;
  workOrderNumber: string;
  type: string;
  diameter: string;
  coupling: string;
  workOrderDate: string;
  transport: string;
  batchCount: number;
  arrivedCount: number;
  inspectionDoneCount: number;
  raw: ParsedRow;
}

interface BatchEntry {
  id: string;
  client: string;
  workOrderNumber: string;
  batchNumber: string;
  qty: string;
  diameter: string;
  pipeFrom: string;
  pipeTo: string;
  class1: string;
  class2: string;
  class3: string;
  repair: string;
  scrapTotal: string;
  statusLabel: string;
  statusKey: StatusKey;
  arrivalDate: string;
  startDate: string;
  endDate: string;
  scrapDetails: {
    rattling: number;
    external: number;
    jetting: number;
    mpi: number;
    drift: number;
    emi: number;
    total: number;
  };
  raw: ParsedRow;
}

const canonicalizeHeader = (header: unknown): string => {
  if (header === undefined || header === null) {
    return '';
  }

  return String(header)
    .trim()
    .toLowerCase()
    .replace(/[\s/]+/g, '_')
    .replace(/[^a-z0-9_]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');
};

const parseSheetData = (sheetData: unknown[]): ParsedSheetData => {
  if (!Array.isArray(sheetData) || sheetData.length === 0) {
    return { headers: [], canonicalHeaders: [], rows: [] };
  }

  if (Array.isArray(sheetData[0])) {
    const headersRow = sheetData[0] as unknown[];
    const canonicalHeaders = headersRow.map(header => canonicalizeHeader(header));
    const rows = sheetData
      .slice(1)
      .filter(row => Array.isArray(row) && row.some(cell => cell !== undefined && cell !== null && cell !== ''))
      .map(row => {
        const typedRow = row as unknown[];
        const data: Record<string, unknown> = {};
        canonicalHeaders.forEach((key, index) => {
          if (!key) return;
          data[key] = typedRow[index];
        });
        return { values: typedRow, data };
      });

    return {
      headers: headersRow.map(header => (header === undefined || header === null ? '' : String(header))),
      canonicalHeaders,
      rows,
    };
  }

  if (typeof sheetData[0] === 'object' && sheetData[0] !== null) {
    const headers = Object.keys(sheetData[0] as Record<string, unknown>);
    const canonicalHeaders = headers.map(header => canonicalizeHeader(header));
    const rows = (sheetData as Record<string, unknown>[]).map(rowObject => {
      const data: Record<string, unknown> = {};
      canonicalHeaders.forEach((key, index) => {
        if (!key) return;
        const header = headers[index];
        data[key] = rowObject[header];
      });
      const values = headers.map(header => rowObject[header]);
      return { values, data };
    });

    return { headers, canonicalHeaders, rows };
  }

  return { headers: [], canonicalHeaders: [], rows: [] };
};

const normalizeString = (value: unknown): string => {
  if (value === undefined || value === null) {
    return '';
  }
  return String(value).trim().toLowerCase();
};

const formatExcelDate = (value: unknown): string => {
  if (value === undefined || value === null || value === '') {
    return '';
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    try {
      const excelEpoch = new Date(1900, 0, 1);
      const date = new Date(excelEpoch.getTime() + (value - 2) * 24 * 60 * 60 * 1000);
      if (!Number.isNaN(date.getTime())) {
        return date.toLocaleDateString();
      }
    } catch (error) {
      console.warn('Failed to format Excel date value:', value, error);
    }
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toLocaleDateString();
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return '';
    }
    const parsed = new Date(trimmed);
    if (!Number.isNaN(parsed.getTime())) {
      return parsed.toLocaleDateString();
    }
    return trimmed;
  }

  return String(value);
};

const getValue = (rowData: Record<string, unknown>, keys: string[], fallback = ''): unknown => {
  for (const key of keys) {
    const canonicalKey = canonicalizeHeader(key);
    if (!canonicalKey) continue;
    const value = rowData[canonicalKey];
    if (value !== undefined && value !== null && value !== '') {
      return value;
    }
  }
  return fallback;
};

const toNumber = (value: unknown): number => {
  if (value === undefined || value === null || value === '') {
    return 0;
  }
  const numberValue = Number(value);
  return Number.isFinite(numberValue) ? numberValue : 0;
};

const getStatusKey = (value: unknown): StatusKey => {
  const normalized = normalizeString(value);
  if (normalized.includes('inspection')) {
    return 'inspection_done';
  }
  if (normalized.includes('arriv')) {
    return 'arrived';
  }
  if (normalized.includes('complete')) {
    return 'completed';
  }
  return normalized ? 'other' : 'other';
};

const parseStoredArray = <T = unknown>(value: string | null, key: string): T[] => {
  if (!value) {
    return [];
  }

  try {
    const parsed = JSON.parse(value);
    if (Array.isArray(parsed)) {
      return parsed as T[];
    }

    console.warn(`Cached ${key} is not an array`);
    return [];
  } catch (error) {
    console.warn(`Error parsing cached ${key}:`, error);
    return [];
  }
};

const SharePointViewer: React.FC = () => {
  const [excelData, setExcelData] = useState<ExcelData>({
    clients: [],
    workOrders: [],
    tubingRegistry: []
  });
  const [selectedClientFilter, setSelectedClientFilter] = useState<string>('all');
  const [workOrderSort, setWorkOrderSort] = useState<'client' | 'batch'>('client');
  const [batchStatusFilter, setBatchStatusFilter] = useState<'all' | 'arrived' | 'inspection_done'>('all');
  const [selectedWorkOrder, setSelectedWorkOrder] = useState<WorkOrderEntry | null>(null);
  const [selectedBatch, setSelectedBatch] = useState<BatchEntry | null>(null);
  const [reportBatch, setReportBatch] = useState<BatchEntry | null>(null);
  const [clientSearchTerm, setClientSearchTerm] = useState('');
  const [workOrderSearchTerm, setWorkOrderSearchTerm] = useState('');
  const [batchSearchTerm, setBatchSearchTerm] = useState('');
  const { toast } = useToast();
  const { isConnected, isConnecting, connect, disconnect, error, cachedClients, cachedWorkOrders } = useSharePoint();
  const navigate = useNavigate();

  useEffect(() => {
    const cachedClientsData = safeLocalStorage.getItem('sharepoint_cached_clients');
    const cachedWorkOrdersData = safeLocalStorage.getItem('sharepoint_cached_workorders');
    const cachedTubingData = safeLocalStorage.getItem('sharepoint_cached_tubing');

    const clients = parseStoredArray<string>(cachedClientsData, 'clients');
    const workOrders = parseStoredArray<unknown>(cachedWorkOrdersData, 'work orders');
    const tubingRegistry = parseStoredArray<unknown>(cachedTubingData, 'tubing');

    setExcelData({ clients, workOrders, tubingRegistry });

  }, []);

  useEffect(() => {
    const cachedTubingData = safeLocalStorage.getItem('sharepoint_cached_tubing');
    const tubingRegistry = parseStoredArray<unknown>(cachedTubingData, 'tubing');

    setExcelData({
      clients: cachedClients,
      workOrders: cachedWorkOrders,
      tubingRegistry
    });
  }, [cachedClients, cachedWorkOrders]);

  const handleConnect = async () => {
    const success = await connect();
    if (success) {
      toast({
        title: "Подключение успешно",
        description: "SharePoint подключен и сохранен",
      });
    } else {
      toast({
        title: "Ошибка подключения",
        description: error || "Не удалось подключиться к SharePoint",
        variant: "destructive",
      });
    }
  };

  const handleDisconnect = () => {
    disconnect();
    setExcelData({ clients: [], workOrders: [], tubingRegistry: [] });
    toast({
      title: "Отключен",
      description: "SharePoint отключен",
    });
  };

  const workOrdersSheet = useMemo(() => parseSheetData(excelData.workOrders), [excelData.workOrders]);
  const tubingSheet = useMemo(() => parseSheetData(excelData.tubingRegistry), [excelData.tubingRegistry]);

  const activeTubingRows = useMemo(() =>
    tubingSheet.rows.filter(row => getStatusKey(getValue(row.data, ['status'])) !== 'completed'),
  [tubingSheet.rows]);

  const workOrderEntries = useMemo<WorkOrderEntry[]>(() => {
    return workOrdersSheet.rows
      .map((row, index) => {
        const client = String(getValue(row.data, ['client', 'client_name'], '')).trim();
        const workOrderNumber = String(getValue(row.data, ['wo_no', 'wo', 'work_order', 'workorder'], '')).trim();

        if (!client && !workOrderNumber) {
          return null;
        }

        const matchingBatches = activeTubingRows.filter(batch =>
          normalizeString(getValue(batch.data, ['client', 'client_name'])) === normalizeString(client) &&
          normalizeString(getValue(batch.data, ['wo_no', 'wo', 'work_order', 'workorder'])) === normalizeString(workOrderNumber)
        );

        const arrivedCount = matchingBatches.filter(batch => getStatusKey(getValue(batch.data, ['status'])) === 'arrived').length;
        const inspectionDoneCount = matchingBatches.filter(batch => getStatusKey(getValue(batch.data, ['status'])) === 'inspection_done').length;

        return {
          id: `${client || 'unknown'}::${workOrderNumber || index}`,
          client,
          workOrderNumber,
          type: String(getValue(row.data, ['type'], '') || ''),
          diameter: String(getValue(row.data, ['diameter'], '') || ''),
          coupling: String(getValue(row.data, ['coupling_replace', 'coupling'], '') || ''),
          workOrderDate: formatExcelDate(getValue(row.data, ['wo_date', 'work_order_date', 'date'])),
          transport: String(getValue(row.data, ['transport', 'transport_company'], '') || ''),
          batchCount: matchingBatches.length,
          arrivedCount,
          inspectionDoneCount,
          raw: row,
        } as WorkOrderEntry;
      })
      .filter((entry): entry is WorkOrderEntry => entry !== null);
  }, [workOrdersSheet.rows, activeTubingRows]);

  const batchEntries = useMemo<BatchEntry[]>(() => {
    return activeTubingRows.map((row, index) => {
      const client = String(getValue(row.data, ['client', 'client_name'], '')).trim();
      const workOrderNumber = String(getValue(row.data, ['wo_no', 'wo', 'work_order', 'workorder'], '')).trim();
      const batchNumber = String(getValue(row.data, ['batch', 'batch_no'], '')).trim();
      const statusValue = getValue(row.data, ['status'], '');
      const statusKey = getStatusKey(statusValue);
      const statusLabel = statusValue ? String(statusValue) : (statusKey === 'arrived' ? 'Arrived' : statusKey === 'inspection_done' ? 'Inspection Done' : 'Unknown');

      const scrapDetails = {
        rattling: toNumber(getValue(row.data, ['rattling_scrap_qty', 'rattling_scrap'])),
        external: toNumber(getValue(row.data, ['external_scrap_qty', 'external_scrap'])),
        jetting: toNumber(getValue(row.data, ['jetting_scrap_qty', 'hydro_scrap_qty', 'jetting_scrap'])),
        mpi: toNumber(getValue(row.data, ['mpi_scrap_qty', 'mpi_scrap'])),
        drift: toNumber(getValue(row.data, ['drift_scrap_qty', 'drift_scrap'])),
        emi: toNumber(getValue(row.data, ['emi_scrap_qty', 'emi_scrap'])),
        total: 0,
      };
      scrapDetails.total = scrapDetails.rattling + scrapDetails.external + scrapDetails.jetting + scrapDetails.mpi + scrapDetails.drift + scrapDetails.emi;

      return {
        id: `${client || 'unknown'}::${workOrderNumber || 'wo'}::${batchNumber || index}`,
        client,
        workOrderNumber,
        batchNumber,
        qty: String(getValue(row.data, ['qty', 'quantity'], '') || ''),
        diameter: String(getValue(row.data, ['diameter'], '') || ''),
        pipeFrom: String(getValue(row.data, ['pipe_from', 'from'], '') || ''),
        pipeTo: String(getValue(row.data, ['pipe_to', 'to'], '') || ''),
        class1: String(getValue(row.data, ['class_1', 'class1'], '') || ''),
        class2: String(getValue(row.data, ['class_2', 'class2'], '') || ''),
        class3: String(getValue(row.data, ['class_3', 'class3'], '') || ''),
        repair: String(getValue(row.data, ['repair'], '') || ''),
        scrapTotal: String(getValue(row.data, ['scrap'], '') || ''),
        statusLabel,
        statusKey,
        arrivalDate: formatExcelDate(getValue(row.data, ['arrival_date', 'arrival'])),
        startDate: formatExcelDate(getValue(row.data, ['start_date', 'inspection_start', 'start'])),
        endDate: formatExcelDate(getValue(row.data, ['end_date', 'inspection_end', 'end'])),
        scrapDetails,
        raw: row,
      } as BatchEntry;
    });
  }, [activeTubingRows]);

  const cachedClientNames = useMemo(() => {
    const values = excelData.clients
      .map(client => {
        if (client === undefined || client === null) return '';
        const name = String(client).trim();
        if (!name) return '';
        if (name.toLowerCase() === 'client') return '';
        return name;
      })
      .filter(Boolean) as string[];

    return Array.from(new Set(values));
  }, [excelData.clients]);

  const clientSummaries = useMemo(() => {
    const map = new Map<string, { workOrders: number; batches: number }>();

    workOrderEntries.forEach(entry => {
      if (!entry.client) return;
      const current = map.get(entry.client) ?? { workOrders: 0, batches: 0 };
      current.workOrders += 1;
      current.batches += entry.batchCount;
      map.set(entry.client, current);
    });

    batchEntries.forEach(entry => {
      if (!entry.client) return;
      const current = map.get(entry.client);
      if (!current) {
        map.set(entry.client, { workOrders: 0, batches: 1 });
        return;
      }

      if (current.workOrders === 0) {
        current.batches += 1;
        map.set(entry.client, current);
      }
    });

    cachedClientNames.forEach(client => {
      if (client && !map.has(client)) {
        map.set(client, { workOrders: 0, batches: 0 });
      }
    });

    return Array.from(map.entries())
      .map(([client, summary]) => ({ client, ...summary }))
      .sort((a, b) => a.client.localeCompare(b.client));
  }, [workOrderEntries, batchEntries, cachedClientNames]);

  const uniqueClients = useMemo(() => clientSummaries.map(summary => summary.client), [clientSummaries]);

  useEffect(() => {
    if (selectedClientFilter !== 'all' && !uniqueClients.includes(selectedClientFilter)) {
      setSelectedClientFilter('all');
    }
  }, [selectedClientFilter, uniqueClients]);

  const filteredClientSummaries = useMemo(() => {
    const term = clientSearchTerm.trim().toLowerCase();
    if (!term) {
      return clientSummaries;
    }
    return clientSummaries.filter(summary => summary.client.toLowerCase().includes(term));
  }, [clientSummaries, clientSearchTerm]);

  const filteredWorkOrders = useMemo(() => {
    const term = workOrderSearchTerm.trim().toLowerCase();
    return workOrderEntries.filter(entry => {
      if (selectedClientFilter !== 'all' && entry.client !== selectedClientFilter) {
        return false;
      }
      if (!term) {
        return true;
      }
      return (entry.workOrderNumber || '').toLowerCase().includes(term);
    });
  }, [selectedClientFilter, workOrderEntries, workOrderSearchTerm]);

  const sortedWorkOrders = useMemo(() => {
    const list = [...filteredWorkOrders];
    if (workOrderSort === 'client') {
      list.sort((a, b) => {
        const clientCompare = a.client.localeCompare(b.client);
        if (clientCompare !== 0) return clientCompare;
        return a.workOrderNumber.localeCompare(b.workOrderNumber);
      });
    } else {
      list.sort((a, b) => {
        if (b.batchCount !== a.batchCount) {
          return b.batchCount - a.batchCount;
        }
        const clientCompare = a.client.localeCompare(b.client);
        if (clientCompare !== 0) return clientCompare;
        return a.workOrderNumber.localeCompare(b.workOrderNumber);
      });
    }
    return list;
  }, [filteredWorkOrders, workOrderSort]);

  const filteredBatches = useMemo(() => {
    const term = batchSearchTerm.trim().toLowerCase();
    return batchEntries.filter(entry => {
      if (selectedClientFilter !== 'all' && entry.client !== selectedClientFilter) {
        return false;
      }
      if (batchStatusFilter === 'all') {
        if (!term) {
          return true;
        }
      } else if (entry.statusKey !== batchStatusFilter) {
        return false;
      }
      if (!term) {
        return true;
      }
      return (entry.batchNumber || '').toLowerCase().includes(term);
    });
  }, [batchEntries, batchStatusFilter, selectedClientFilter, batchSearchTerm]);

  const sortedBatches = useMemo(() => {
    return [...filteredBatches].sort((a, b) => {
      const clientCompare = a.client.localeCompare(b.client);
      if (clientCompare !== 0) return clientCompare;
      const woCompare = a.workOrderNumber.localeCompare(b.workOrderNumber);
      if (woCompare !== 0) return woCompare;
      return a.batchNumber.localeCompare(b.batchNumber);
    });
  }, [filteredBatches]);

  const getStatusBadgeClass = (statusKey: StatusKey) => {
    switch (statusKey) {
      case 'arrived':
        return 'bg-blue-100 text-blue-700 border border-blue-200';
      case 'inspection_done':
        return 'bg-emerald-100 text-emerald-700 border border-emerald-200';
      default:
        return 'bg-slate-100 text-slate-700 border border-slate-200';
    }
  };

  const totalClients = clientSummaries.length;
  const totalWorkOrders = workOrderEntries.length;
  const totalBatches = batchEntries.length;

  if (!isConnected) {
    return (
      <div className="container mx-auto p-6">
        <div className="mb-6">
          <Button
            onClick={() => navigate('/')}
            variant="outline"
            className="mb-4 border-2 hover:bg-gray-50"
          >
            <ArrowLeft className="mr-2 h-4 w-4" />
            Назад в меню
          </Button>
        </div>
        <Card className="max-w-md mx-auto border-2 shadow-lg">
          <CardHeader className="text-center">
            <div className="mx-auto mb-4 flex h-12 w-12 items-center justify-center rounded-full bg-blue-100 border-2 border-blue-200">
              <FileSpreadsheet className="h-6 w-6 text-blue-600" />
            </div>
            <CardTitle className="text-xl">Dashboard</CardTitle>
            <CardDescription className="text-base">
              Connect to Microsoft Graph to open the SharePoint dashboard
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <Button
              onClick={handleConnect}
              disabled={isConnecting}
              className="w-full h-12 text-base font-semibold"
            >
              {isConnecting && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}
              Sign in with Microsoft
            </Button>
            <p className="text-xs text-muted-foreground text-center">
              Requires access to kzprimeestate.sharepoint.com
            </p>
          </CardContent>
        </Card>
      </div>
    );
  }

  return (
    <div className="container mx-auto p-6 space-y-6">
      <div className="mb-6">
        <Button
          onClick={() => navigate('/')}
          variant="outline"
          className="mb-4 border-2 hover:bg-gray-50"
        >
          <ArrowLeft className="mr-2 h-4 w-4" />
          Назад в меню
        </Button>
      </div>

      <div className="flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
        <div>
          <h1 className="text-3xl font-bold">Dashboard</h1>
          <p className="text-muted-foreground text-lg">
            Overview of active clients, work orders, and tubing batches from SharePoint Excel
          </p>
        </div>
        <div className="flex gap-3">
          <Button onClick={handleDisconnect} variant="outline" className="border-2 hover:bg-red-50">
            Sign Out
          </Button>
        </div>
      </div>

      <div className="grid grid-cols-1 gap-4 md:grid-cols-3">
        <Card className="border-2 shadow-lg hover:shadow-xl transition-shadow">
          <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-3">
            <CardTitle className="text-base font-semibold">Clients</CardTitle>
            <Users className="h-5 w-5 text-blue-600" />
          </CardHeader>
          <CardContent>
            <div className="text-3xl font-bold text-blue-600">{totalClients}</div>
            <p className="text-sm text-muted-foreground">Unique clients with active records</p>
          </CardContent>
        </Card>
        <Card className="border-2 shadow-lg hover:shadow-xl transition-shadow">
          <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-3">
            <CardTitle className="text-base font-semibold">Work Order Info</CardTitle>
            <Briefcase className="h-5 w-5 text-green-600" />
          </CardHeader>
          <CardContent>
            <div className="text-3xl font-bold text-green-600">{totalWorkOrders}</div>
            <p className="text-sm text-muted-foreground">Active Work Orders</p>
          </CardContent>
        </Card>
        <Card className="border-2 shadow-lg hover:shadow-xl transition-shadow">
          <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-3">
            <CardTitle className="text-base font-semibold">Batch Info</CardTitle>
            <Database className="h-5 w-5 text-purple-600" />
          </CardHeader>
          <CardContent>
            <div className="text-3xl font-bold text-purple-600">{totalBatches}</div>
            <p className="text-sm text-muted-foreground">Batches in Arrived or Inspection Done status</p>
          </CardContent>
        </Card>
      </div>

      <Tabs defaultValue="clients" className="space-y-6">
        <TabsList className="grid w-full grid-cols-3 border-2">
          <TabsTrigger value="clients" className="font-semibold">Clients</TabsTrigger>
          <TabsTrigger value="workorders" className="font-semibold">Work Order Info</TabsTrigger>
          <TabsTrigger value="batches" className="font-semibold">Batch Info</TabsTrigger>
        </TabsList>

        <TabsContent value="clients" className="space-y-6">
          <Card className="border-2 shadow-lg">
            <CardHeader>
              <CardTitle className="flex items-center gap-2 text-xl">
                <Users className="h-6 w-6 text-blue-600" />
                Clients Overview
              </CardTitle>
              <CardDescription className="text-base">
                Unique clients with the number of active work orders and batches
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="space-y-2 md:text-right">
                <div className="text-sm text-muted-foreground">Search clients</div>
                <Input
                  value={clientSearchTerm}
                  onChange={event => setClientSearchTerm(event.target.value)}
                  placeholder="Search by client name"
                  className="w-full md:w-[260px]"
                />
              </div>

              {filteredClientSummaries.length > 0 ? (
                <div className="space-y-3">
                  {filteredClientSummaries.map(summary => (
                    <div
                      key={summary.client}
                      className="flex flex-col gap-3 rounded-lg border px-4 py-3 shadow-sm md:flex-row md:items-center md:justify-between"
                    >
                      <div>
                        <p className="text-xs uppercase text-muted-foreground">Client</p>
                        <p className="text-lg font-semibold text-slate-900">{summary.client}</p>
                      </div>
                      <div className="flex flex-col gap-2 text-right md:flex-row md:gap-6 md:text-left">
                        <div>
                          <p className="text-xs uppercase text-muted-foreground">Work Orders</p>
                          <p className="text-lg font-semibold text-slate-900">{summary.workOrders}</p>
                        </div>
                        <div>
                          <p className="text-xs uppercase text-muted-foreground">Active Batches</p>
                          <p className="text-lg font-semibold text-slate-900">{summary.batches}</p>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              ) : clientSummaries.length > 0 ? (
                <div className="text-center py-10 text-muted-foreground">
                  No clients match the search term.
                </div>
              ) : (
                <div className="text-center py-10 text-muted-foreground">
                  No client data available. Use the main dashboard to sync SharePoint data.
                </div>
              )}
            </CardContent>
          </Card>
        </TabsContent>

        <TabsContent value="workorders" className="space-y-6">
          <Card className="border-2 shadow-lg">
            <CardHeader>
              <CardTitle className="flex items-center gap-2 text-xl">
                <Briefcase className="h-6 w-6 text-green-600" />
                Work Order Info
              </CardTitle>
              <CardDescription className="text-base">
                Sort work orders by client or by batch count.
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-6">
              <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
                <div className="space-y-2">
                  <div className="flex items-center gap-2 text-sm text-muted-foreground">
                    <Filter className="h-4 w-4" />
                    <span>Sort work orders</span>
                  </div>
                  <ToggleGroup
                    type="single"
                    value={workOrderSort}
                    onValueChange={value => value && setWorkOrderSort(value as 'client' | 'batch')}
                    className="rounded-md border bg-muted/40 p-1 md:inline-flex"
                  >
                    <ToggleGroupItem value="client" className="flex-1">By client</ToggleGroupItem>
                    <ToggleGroupItem value="batch" className="flex-1">By batch count</ToggleGroupItem>
                  </ToggleGroup>
                </div>
                <div className="space-y-2 md:text-right">
                  <div className="text-sm text-muted-foreground">Filter by client</div>
                  <Select value={selectedClientFilter} onValueChange={setSelectedClientFilter}>
                    <SelectTrigger className="w-full md:w-[220px]">
                      <SelectValue placeholder="All clients" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="all">All clients</SelectItem>
                      {uniqueClients.map(client => (
                        <SelectItem key={client} value={client}>
                          {client}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-2 md:text-right">
                  <div className="text-sm text-muted-foreground">Search work orders</div>
                  <Input
                    value={workOrderSearchTerm}
                    onChange={event => setWorkOrderSearchTerm(event.target.value)}
                    placeholder="Search by work order number"
                    className="w-full md:w-[220px]"
                  />
                </div>
              </div>

              <div className="overflow-hidden rounded-lg border">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead className="w-[140px]">Work Order</TableHead>
                      <TableHead>Client</TableHead>
                      <TableHead>Type</TableHead>
                      <TableHead>Diameter</TableHead>
                      <TableHead className="text-center">Batches</TableHead>
                      <TableHead className="text-center">Arrived</TableHead>
                      <TableHead className="text-center">Inspection Done</TableHead>
                      <TableHead className="w-[120px] text-right">Actions</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {sortedWorkOrders.length > 0 ? (
                      sortedWorkOrders.map(entry => (
                        <TableRow key={entry.id}>
                          <TableCell className="font-medium">{entry.workOrderNumber || '—'}</TableCell>
                          <TableCell>{entry.client || '—'}</TableCell>
                          <TableCell>{entry.type || '—'}</TableCell>
                          <TableCell>{entry.diameter || '—'}</TableCell>
                          <TableCell className="text-center font-semibold">{entry.batchCount}</TableCell>
                          <TableCell className="text-center">{entry.arrivedCount}</TableCell>
                          <TableCell className="text-center">{entry.inspectionDoneCount}</TableCell>
                          <TableCell className="text-right">
                            <Button variant="outline" size="sm" onClick={() => setSelectedWorkOrder(entry)}>
                              <Info className="mr-2 h-4 w-4" /> Info
                            </Button>
                          </TableCell>
                        </TableRow>
                      ))
                    ) : (
                      <TableRow>
                        <TableCell colSpan={8} className="py-6 text-center text-muted-foreground">
                          No active work orders available for the selected filters.
                        </TableCell>
                      </TableRow>
                    )}
                  </TableBody>
                </Table>
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        <TabsContent value="batches" className="space-y-6">
          <Card className="border-2 shadow-lg">
            <CardHeader>
              <CardTitle className="flex items-center gap-2 text-xl">
                <Database className="h-6 w-6 text-purple-600" />
                Batch Info
              </CardTitle>
              <CardDescription className="text-base">
                View Arrived and Inspection Done batches.
              </CardDescription>
            </CardHeader>
            <CardContent className="space-y-6">
              <div className="flex flex-col gap-4 md:flex-row md:items-end md:justify-between">
                <div className="space-y-2">
                  <div className="text-sm text-muted-foreground">Show batches by status</div>
                  <ToggleGroup
                    type="single"
                    value={batchStatusFilter}
                    onValueChange={value => value && setBatchStatusFilter(value as 'all' | 'arrived' | 'inspection_done')}
                    className="rounded-md border bg-muted/40 p-1 md:inline-flex"
                  >
                    <ToggleGroupItem value="all" className="flex-1">All</ToggleGroupItem>
                    <ToggleGroupItem value="arrived" className="flex-1">Arrived</ToggleGroupItem>
                    <ToggleGroupItem value="inspection_done" className="flex-1">Inspection Done</ToggleGroupItem>
                  </ToggleGroup>
                </div>
                <div className="space-y-2 md:text-right">
                  <div className="text-sm text-muted-foreground">Search batches</div>
                  <Input
                    value={batchSearchTerm}
                    onChange={event => setBatchSearchTerm(event.target.value)}
                    placeholder="Search by batch number"
                    className="w-full md:w-[220px]"
                  />
                </div>
              </div>

              <div className="overflow-hidden rounded-lg border">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead className="w-[120px]">Batch</TableHead>
                      <TableHead>Work Order</TableHead>
                      <TableHead>Client</TableHead>
                      <TableHead className="text-center">Qty</TableHead>
                      <TableHead>Diameter</TableHead>
                      <TableHead>Status</TableHead>
                      <TableHead className="w-[120px] text-right">Actions</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {sortedBatches.length > 0 ? (
                      sortedBatches.map(entry => (
                        <TableRow key={entry.id}>
                          <TableCell className="font-medium">{entry.batchNumber || '—'}</TableCell>
                          <TableCell>{entry.workOrderNumber || '—'}</TableCell>
                          <TableCell>{entry.client || '—'}</TableCell>
                          <TableCell className="text-center">{entry.qty || '—'}</TableCell>
                          <TableCell>{entry.diameter || '—'}</TableCell>
                          <TableCell>
                            <Badge className={getStatusBadgeClass(entry.statusKey)}>{entry.statusLabel}</Badge>
                          </TableCell>
                          <TableCell className="text-right">
                            <Button variant="outline" size="sm" onClick={() => setSelectedBatch(entry)}>
                              <Info className="mr-2 h-4 w-4" /> Info
                            </Button>
                          </TableCell>
                        </TableRow>
                      ))
                    ) : (
                      <TableRow>
                        <TableCell colSpan={7} className="py-6 text-center text-muted-foreground">
                          No batches match the selected filters.
                        </TableCell>
                      </TableRow>
                    )}
                  </TableBody>
                </Table>
              </div>
            </CardContent>
          </Card>
        </TabsContent>
      </Tabs>

      <Dialog open={!!selectedWorkOrder} onOpenChange={open => !open && setSelectedWorkOrder(null)}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Work Order Info</DialogTitle>
            <DialogDescription>Details for the selected work order</DialogDescription>
          </DialogHeader>
          {selectedWorkOrder && (
            <div className="space-y-4">
              <div className="grid gap-3 sm:grid-cols-2">
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Client</p>
                  <p className="font-semibold text-slate-900">{selectedWorkOrder.client || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Work Order Number</p>
                  <p className="font-semibold text-slate-900">{selectedWorkOrder.workOrderNumber || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Type</p>
                  <p className="font-semibold text-slate-900">{selectedWorkOrder.type || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Diameter</p>
                  <p className="font-semibold text-slate-900">{selectedWorkOrder.diameter || '—'}</p>
                </div>
              </div>
              <div className="grid gap-3 sm:grid-cols-2">
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Coupling Replace</p>
                  <p className="font-semibold text-slate-900">{selectedWorkOrder.coupling || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Work Order Date</p>
                  <p className="font-semibold text-slate-900">{selectedWorkOrder.workOrderDate || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Transport</p>
                  <p className="font-semibold text-slate-900">{selectedWorkOrder.transport || '—'}</p>
                </div>
              </div>
              <div className="grid gap-3 sm:grid-cols-3">
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Total Batches</p>
                  <p className="font-semibold text-slate-900">{selectedWorkOrder.batchCount}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Arrived</p>
                  <p className="font-semibold text-slate-900">{selectedWorkOrder.arrivedCount}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Inspection Done</p>
                  <p className="font-semibold text-slate-900">{selectedWorkOrder.inspectionDoneCount}</p>
                </div>
              </div>
            </div>
          )}
        </DialogContent>
      </Dialog>

      <Dialog open={!!selectedBatch} onOpenChange={open => !open && setSelectedBatch(null)}>
        <DialogContent className="max-w-2xl">
          <DialogHeader>
            <DialogTitle>Batch Info</DialogTitle>
            <DialogDescription>Details for the selected batch</DialogDescription>
          </DialogHeader>
          {selectedBatch && (
            <div className="space-y-4">
              <div className="grid gap-3 sm:grid-cols-2">
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Client</p>
                  <p className="font-semibold text-slate-900">{selectedBatch.client || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Work Order</p>
                  <p className="font-semibold text-slate-900">{selectedBatch.workOrderNumber || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Batch</p>
                  <p className="font-semibold text-slate-900">{selectedBatch.batchNumber || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Quantity</p>
                  <p className="font-semibold text-slate-900">{selectedBatch.qty || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Diameter</p>
                  <p className="font-semibold text-slate-900">{selectedBatch.diameter || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Status</p>
                  <Badge className={getStatusBadgeClass(selectedBatch.statusKey)}>{selectedBatch.statusLabel}</Badge>
                </div>
              </div>
              <div className="grid gap-3 sm:grid-cols-3">
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Pipe From</p>
                  <p className="font-semibold text-slate-900">{selectedBatch.pipeFrom || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Pipe To</p>
                  <p className="font-semibold text-slate-900">{selectedBatch.pipeTo || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Scrap</p>
                  <p className="font-semibold text-slate-900">{selectedBatch.scrapTotal || '—'}</p>
                </div>
              </div>
              <div className="grid gap-3 sm:grid-cols-4">
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Class 1</p>
                  <p className="font-semibold text-slate-900">{selectedBatch.class1 || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Class 2</p>
                  <p className="font-semibold text-slate-900">{selectedBatch.class2 || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Class 3</p>
                  <p className="font-semibold text-slate-900">{selectedBatch.class3 || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Repair</p>
                  <p className="font-semibold text-slate-900">{selectedBatch.repair || '—'}</p>
                </div>
              </div>
              {selectedBatch.statusKey === 'arrived' && (
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Arrival Date</p>
                  <p className="font-semibold text-slate-900">{selectedBatch.arrivalDate || '—'}</p>
                </div>
              )}
              {selectedBatch.statusKey === 'inspection_done' && (
                <div className="grid gap-3 sm:grid-cols-2">
                  <div>
                    <p className="text-xs uppercase text-muted-foreground">Start Date</p>
                    <p className="font-semibold text-slate-900">{selectedBatch.startDate || '—'}</p>
                  </div>
                  <div>
                    <p className="text-xs uppercase text-muted-foreground">End Date</p>
                    <p className="font-semibold text-slate-900">{selectedBatch.endDate || '—'}</p>
                  </div>
                </div>
              )}
              <div className="flex flex-col gap-3 border-t pt-3 sm:flex-row sm:items-center sm:justify-between">
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Inspection Report</p>
                  <p className="text-sm text-muted-foreground">
                    View scrap quantities recorded for this batch when inspection is done.
                  </p>
                </div>
                <div className="flex gap-3">
                  {selectedBatch.statusKey === 'inspection_done' ? (
                    <Button variant="outline" onClick={() => setReportBatch(selectedBatch)}>
                      View Report
                    </Button>
                  ) : (
                    <Button variant="outline" disabled>
                      Available after inspection
                    </Button>
                  )}
                </div>
              </div>
            </div>
          )}
        </DialogContent>
      </Dialog>

      <Dialog open={!!reportBatch} onOpenChange={open => !open && setReportBatch(null)}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Inspection Report</DialogTitle>
            <DialogDescription>
              Detailed scrap quantities for batch {reportBatch?.batchNumber || ''}
            </DialogDescription>
          </DialogHeader>
          {reportBatch && (
            <div className="space-y-4">
              <div className="grid gap-3 sm:grid-cols-2">
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Client</p>
                  <p className="font-semibold text-slate-900">{reportBatch.client || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Work Order</p>
                  <p className="font-semibold text-slate-900">{reportBatch.workOrderNumber || '—'}</p>
                </div>
              </div>
              <div className="grid gap-3 sm:grid-cols-2">
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Batch</p>
                  <p className="font-semibold text-slate-900">{reportBatch.batchNumber || '—'}</p>
                </div>
                <div>
                  <p className="text-xs uppercase text-muted-foreground">Total Scrap</p>
                  <p className="font-semibold text-slate-900">{reportBatch.scrapDetails.total}</p>
                </div>
              </div>
              <div className="grid gap-3 sm:grid-cols-2">
                <div className="rounded-lg border p-3 shadow-sm">
                  <p className="text-xs uppercase text-muted-foreground">Rattling Scrap Qty</p>
                  <p className="text-lg font-semibold text-slate-900">{reportBatch.scrapDetails.rattling}</p>
                </div>
                <div className="rounded-lg border p-3 shadow-sm">
                  <p className="text-xs uppercase text-muted-foreground">External Scrap Qty</p>
                  <p className="text-lg font-semibold text-slate-900">{reportBatch.scrapDetails.external}</p>
                </div>
                <div className="rounded-lg border p-3 shadow-sm">
                  <p className="text-xs uppercase text-muted-foreground">Jetting Scrap Qty</p>
                  <p className="text-lg font-semibold text-slate-900">{reportBatch.scrapDetails.jetting}</p>
                </div>
                <div className="rounded-lg border p-3 shadow-sm">
                  <p className="text-xs uppercase text-muted-foreground">MPI Scrap Qty</p>
                  <p className="text-lg font-semibold text-slate-900">{reportBatch.scrapDetails.mpi}</p>
                </div>
                <div className="rounded-lg border p-3 shadow-sm">
                  <p className="text-xs uppercase text-muted-foreground">Drift Scrap Qty</p>
                  <p className="text-lg font-semibold text-slate-900">{reportBatch.scrapDetails.drift}</p>
                </div>
                <div className="rounded-lg border p-3 shadow-sm">
                  <p className="text-xs uppercase text-muted-foreground">EMI Scrap Qty</p>
                  <p className="text-lg font-semibold text-slate-900">{reportBatch.scrapDetails.emi}</p>
                </div>
              </div>
            </div>
          )}
        </DialogContent>
      </Dialog>
    </div>
  );
};

export default SharePointViewer;
