import React, { useState, useEffect } from 'react';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Badge } from "@/components/ui/badge";
import { Loader2, FileSpreadsheet, Users, Briefcase, Database, ArrowLeft } from "lucide-react";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useToast } from "@/hooks/use-toast";
import { useNavigate } from "react-router-dom";

interface ExcelData {
  clients: string[];
  workOrders: any[];
  tubingRegistry: any[];
}

const parseStoredArray = (value: string | null, key: string) => {
  if (!value) {
    return [];
  }

  try {
    const parsed = JSON.parse(value);
    if (Array.isArray(parsed)) {
      return parsed;
    }

    console.warn(`Cached ${key} is not an array`);
    return [];
  } catch (error) {
    console.warn(`Error parsing cached ${key}:`, error);
    return [];
  }
};

const SharePointViewer: React.FC = () => {
  const [isLoading, setIsLoading] = useState(false);
  const [excelData, setExcelData] = useState<ExcelData>({
    clients: [],
    workOrders: [],
    tubingRegistry: []
  });
  const [dataLoaded, setDataLoaded] = useState(false);
  const { toast } = useToast();
  const { isConnected, isConnecting, sharePointService, connect, disconnect, error, cachedClients, cachedWorkOrders } = useSharePoint();
  const navigate = useNavigate();

  // Загружаем кешированные данные сразу при монтировании
  useEffect(() => {
    loadCachedData();
  }, []); // Загружаем только один раз при монтировании

  // Убираем автоматическую загрузку - данные управляются централизованно

  // Убираем автоматическое обновление - теперь только через главную страницу

  const loadCachedData = () => {
    try {
      const cachedClientsData = localStorage.getItem('sharepoint_cached_clients');
      const cachedWorkOrdersData = localStorage.getItem('sharepoint_cached_workorders');
      const cachedTubingData = localStorage.getItem('sharepoint_cached_tubing');

      const clients = parseStoredArray(cachedClientsData, 'clients');
      const workOrders = parseStoredArray(cachedWorkOrdersData, 'work orders');
      const tubingRegistry = parseStoredArray(cachedTubingData, 'tubing');

      setExcelData({
        clients,
        workOrders,
        tubingRegistry
      });

      if (clients.length > 0) {
        console.log('📦 Viewer loaded cached clients:', clients.length);
      }
      if (workOrders.length > 0) {
        console.log('📦 Viewer loaded cached work orders:', workOrders.length);
      }
      if (tubingRegistry.length > 0) {
        console.log('📦 Viewer loaded cached tubing registry:', tubingRegistry.length);
      }

      const hasData = clients.length > 0 || workOrders.length > 0 || tubingRegistry.length > 0;
      setDataLoaded(hasData);
      if (hasData) {
        console.log('✅ SharePoint Viewer ready with cached data');
      }
    } catch (error) {
      console.error('Error loading cached data:', error);
    }
  };

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
    setDataLoaded(false);
    toast({
      title: "Отключен",
      description: "SharePoint отключен",
    });
  };

  // Мгновенное обновление данных при изменении кеша
  useEffect(() => {
    const cachedTubingData = localStorage.getItem('sharepoint_cached_tubing');
    const tubingRegistry = parseStoredArray(cachedTubingData, 'tubing');

    setExcelData({
      clients: cachedClients,
      workOrders: cachedWorkOrders,
      tubingRegistry
    });

    const hasData = cachedClients.length > 0 || cachedWorkOrders.length > 0 || tubingRegistry.length > 0;
    setDataLoaded(hasData);
  }, [cachedClients, cachedWorkOrders]);

  const renderTable = (data: any[], title: string, icon: React.ReactNode) => {
    if (!data || data.length === 0) {
      return (
        <div className="text-center py-8 text-muted-foreground">
          No {title.toLowerCase()} data available
        </div>
      );
    }

    // Debug the data structure
    console.log(`🔍 ${title} data structure:`, data);
    console.log(`🔍 First item:`, data[0]);
    console.log(`🔍 Is first item array:`, Array.isArray(data[0]));

    // Handle both array of arrays and array of objects
    let headers: string[] = [];
    let rows: any[][] = [];

    if (Array.isArray(data[0])) {
      // Data is array of arrays (Excel format)
      headers = data[0] || [];
      rows = data.slice(1).filter(row => Array.isArray(row) && row.length > 0);
    } else if (typeof data[0] === 'object' && data[0] !== null) {
      // Data is array of objects
      headers = Object.keys(data[0]);
      rows = data.map(item => headers.map(header => item[header] || ''));
    } else {
      // Unknown format
      console.warn(`Unknown data format for ${title}:`, data);
      return (
        <div className="text-center py-8 text-muted-foreground">
          Invalid data format for {title.toLowerCase()}
        </div>
      );
    }

    return (
      <div className="space-y-4">
        <div className="flex items-center gap-2">
          {icon}
          <h3 className="text-lg font-semibold">{title}</h3>
          <Badge variant="secondary">{rows.length} records</Badge>
        </div>
        <div className="border-2 rounded-lg shadow-sm">
          <Table>
            <TableHeader>
              <TableRow className="bg-gray-50">
                {headers.map((header: string, index: number) => (
                  <TableHead key={index} className="font-bold text-gray-700 border-r last:border-r-0">
                    {header}
                  </TableHead>
                ))}
              </TableRow>
            </TableHeader>
            <TableBody>
              {rows.map((row: any[], rowIndex: number) => (
                <TableRow key={rowIndex} className="hover:bg-gray-50 border-b">
                  {headers.map((_: string, cellIndex: number) => {
                    const header = headers[cellIndex]?.toLowerCase() || '';
                    const cell = Array.isArray(row) ? row[cellIndex] : undefined;
                    let displayValue = (cell !== undefined && cell !== null) ? cell.toString() : '';
                    
                    // Форматируем даты для полей Start_Date и End_Date
                    if ((header.includes('date') || header.includes('_date')) && cell && typeof cell === 'number') {
                      try {
                        // Excel хранит даты как число дней с 1 января 1900
                        const excelEpoch = new Date(1900, 0, 1);
                        const date = new Date(excelEpoch.getTime() + (cell - 2) * 24 * 60 * 60 * 1000);
                        displayValue = date.toLocaleDateString('ru-RU');
                      } catch (error) {
                        console.warn('Error formatting date:', cell, error);
                      }
                    }
                    
                    return (
                      <TableCell key={cellIndex} className="max-w-[200px] truncate border-r last:border-r-0 py-3">
                        {displayValue}
                      </TableCell>
                    );
                  })}
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </div>
      </div>
    );
  };

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
            <CardTitle className="text-xl">SharePoint Excel Viewer</CardTitle>
            <CardDescription className="text-base">
              Connect to Microsoft Graph to view pipe inspection data from SharePoint
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
      
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-3xl font-bold">SharePoint Excel Viewer</h1>
          <p className="text-muted-foreground text-lg">
            View and manage pipe inspection data from pipe_inspection.xlsm
          </p>
        </div>
        <div className="flex gap-3">
          <Button onClick={handleDisconnect} variant="outline" className="border-2 hover:bg-red-50">
            Sign Out
          </Button>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <Card className="border-2 shadow-lg hover:shadow-xl transition-shadow">
          <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-3">
            <CardTitle className="text-base font-semibold">Clients</CardTitle>
            <Users className="h-5 w-5 text-blue-600" />
          </CardHeader>
          <CardContent>
            <div className="text-3xl font-bold text-blue-600">{excelData.clients.filter(client => client && client.trim()).length}</div>
          </CardContent>
        </Card>
        <Card className="border-2 shadow-lg hover:shadow-xl transition-shadow">
          <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-3">
            <CardTitle className="text-base font-semibold">Work Orders</CardTitle>
            <Briefcase className="h-5 w-5 text-green-600" />
          </CardHeader>
          <CardContent>
            <div className="text-3xl font-bold text-green-600">{Math.max(0, excelData.workOrders.length - 1)}</div>
          </CardContent>
        </Card>
        <Card className="border-2 shadow-lg hover:shadow-xl transition-shadow">
          <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-3">
            <CardTitle className="text-base font-semibold">Tubing Records</CardTitle>
            <Database className="h-5 w-5 text-purple-600" />
          </CardHeader>
          <CardContent>
            <div className="text-3xl font-bold text-purple-600">{Math.max(0, excelData.tubingRegistry.length - 1)}</div>
          </CardContent>
        </Card>
      </div>

      <Tabs defaultValue="clients" className="space-y-6">
        <TabsList className="grid w-full grid-cols-3 border-2">
          <TabsTrigger value="clients" className="font-semibold">Clients</TabsTrigger>
          <TabsTrigger value="workorders" className="font-semibold">Work Orders</TabsTrigger>
          <TabsTrigger value="tubing" className="font-semibold">Tubing Registry</TabsTrigger>
        </TabsList>
        
        <TabsContent value="clients" className="space-y-6">
          <Card className="border-2 shadow-lg">
            <CardHeader>
              <CardTitle className="flex items-center gap-2 text-xl">
                <Users className="h-6 w-6 text-blue-600" />
                Clients List
              </CardTitle>
              <CardDescription className="text-base">
                Available clients from SharePoint Excel file
              </CardDescription>
            </CardHeader>
            <CardContent>
              {excelData.clients.filter(client => client && client.trim()).length > 0 ? (
                <div className="grid grid-cols-1 md:grid-cols-3 gap-2">
                  {excelData.clients.filter(client => client && client.trim()).map((client, index) => (
                    <Badge key={index} variant="outline" className="justify-center p-3 text-sm font-medium border-2 hover:bg-blue-50">
                      {client}
                    </Badge>
                  ))}
                </div>
              ) : (
                <div className="text-center py-8 text-muted-foreground">
                  No clients loaded. Use "Load Data" button on main page to load from SharePoint.
                </div>
              )}
            </CardContent>
          </Card>
        </TabsContent>
        
        <TabsContent value="workorders" className="space-y-6">
          <Card className="border-2 shadow-lg">
            <CardContent className="pt-6">
              {renderTable(excelData.workOrders, "Work Orders", <Briefcase className="h-5 w-5 text-green-600" />)}
            </CardContent>
          </Card>
        </TabsContent>
        
        <TabsContent value="tubing" className="space-y-6">
          <Card className="border-2 shadow-lg">
            <CardContent className="pt-6">
              {renderTable(excelData.tubingRegistry, "Tubing Registry", <Database className="h-5 w-5 text-purple-600" />)}
            </CardContent>
          </Card>
        </TabsContent>
      </Tabs>
    </div>
  );
};

export default SharePointViewer;
