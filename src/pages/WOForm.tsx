import { useState, useEffect } from "react";
import { Header } from "@/components/layout/Header";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { ArrowLeft } from "lucide-react";
import { useNavigate } from "react-router-dom";
import { useAuth } from "@/contexts/AuthContext";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useToast } from "@/hooks/use-toast";
import { useSharePointInstantData } from "@/hooks/useInstantData";
import { DateInputField } from "@/components/ui/date-input";
import { safeLocalStorage } from '@/lib/safe-storage';

export default function WOForm() {
  const navigate = useNavigate();
  const { user } = useAuth();
  const { sharePointService, isConnected, refreshDataInBackground } = useSharePoint();
  const { toast } = useToast();
  const { clients } = useSharePointInstantData();
  const [availableClients, setAvailableClients] = useState<string[]>(clients);
  const [existingWorkOrders, setExistingWorkOrders] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(false);

  // Core fields
  const [client, setClient] = useState("");
  const [woNo, setWoNo] = useState("");
  const [woDate, setWoDate] = useState("");
  const [typeOfWO, setTypeOfWO] = useState<'OCTG' | 'Coupling'>('OCTG');
  const [pipeType, setPipeType] = useState<'Tubing' | 'Sucker Rod'>('Tubing');
  const [diameter, setDiameter] = useState("");
  const [plannedQty, setPlannedQty] = useState("");
  const [priceType, setPriceType] = useState<'Fixed' | 'Stage Based' | ''>('');
  const [price, setPrice] = useState(""); // Fixed or Coupling Replace
  const [transport, setTransport] = useState<'Client' | 'TCC' | ''>('');
  const [transportCost, setTransportCost] = useState("");

  // Stage-based prices
  const [stagePrices, setStagePrices] = useState({
    rattling_price: "",
    external_price: "",
    hydro_price: "",
    mpi_price: "",
    drift_price: "",
    emi_price: "",
    marking_price: "",
  });

  const keyCol = woNo && client && (typeOfWO === 'Coupling' ? 'Tubing' : pipeType) && (typeOfWO === 'Coupling' ? '' : diameter)
    ? `${woNo} - ${client} - ${(typeOfWO === 'Coupling' ? 'Tubing' : pipeType)} - ${(typeOfWO === 'Coupling' ? '' : diameter)}`.replace(/\s+-\s+$/,'')
    : "";

  // Мгновенное обновление клиентов из кеша - данные всегда доступны
  useEffect(() => {
    const filteredClients = clients.filter(client => client && client.trim());
    if (filteredClients.length > 0) {
      setAvailableClients(filteredClients);
      console.log('⚡ WOForm loaded clients from cache:', filteredClients.length);
    }
  }, [clients]);

  // Load existing work orders when client is selected
  useEffect(() => {
    const loadWorkOrders = async () => {
      if (!client || !sharePointService) {
        setExistingWorkOrders([]);
        return;
      }
      
      try {
        console.log(`🔍 Loading work orders for client: ${client}`);
        const workOrders = await sharePointService.getWorkOrdersByClient(client);
        console.log(`📋 Found ${workOrders.length} existing work orders for ${client}:`, workOrders);
        setExistingWorkOrders(workOrders);
      } catch (error) {
        console.error('❌ Error loading work orders:', error);
        setExistingWorkOrders([]);
      }
    };
    loadWorkOrders();
  }, [client, sharePointService]);

  // Helpers to enforce numeric formats
  const sanitizeDecimal = (value: string) => {
    let v = value.replace(/,/g, '.');
    v = v.replace(/[^0-9.]/g, '');
    const parts = v.split('.');
    if (parts.length > 2) {
      v = parts[0] + '.' + parts.slice(1).join('');
    }
    return v;
  };
  const sanitizeInteger = (value: string) => value.replace(/[^0-9]/g, '');

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!user) {
      toast({
        title: "Ошибка",
        description: "Пожалуйста, войдите в систему",
        variant: "destructive",
      });
      return;
    }

    if (!sharePointService) {
      toast({
        title: "Ошибка",
        description: "SharePoint не подключен",
        variant: "destructive",
      });
      return;
    }

    // Validate required fields
    if (!woNo || !client) {
      toast({
        title: "Ошибка валидации",
        description: "Заполните номер Work Order и выберите клиента",
        variant: "destructive",
      });
      return;
    }

    // Check if work order already exists for this client
    console.log(`🔍 Checking duplicate: WO ${woNo} for client ${client}`);
    console.log(`📋 Existing work orders for ${client}:`, existingWorkOrders);
    
    const isDuplicate = existingWorkOrders.some(wo => wo.toString().trim() === woNo.toString().trim());
    console.log(`🔍 Duplicate check result:`, isDuplicate);
    
    if (isDuplicate) {
      console.log(`❌ DUPLICATE FOUND: Work Order ${woNo} already exists for client ${client}`);
        toast({
          title: "🚫 Work Order уже существует",
          description: (
            <div className="space-y-2">
              <p className="font-bold text-white">
                Work Order <span className="bg-white text-red-600 px-2 py-1 rounded font-mono font-bold">{woNo}</span> уже существует для клиента <span className="bg-white text-blue-600 px-2 py-1 rounded font-bold">{client}</span>
              </p>
              <p className="text-sm text-white font-medium">
                💡 Пожалуйста, выберите другой номер Work Order
              </p>
            </div>
          ),
          variant: "destructive",
          duration: 8000,
        });
      return;
    }
    
    console.log(`✅ Work Order ${woNo} is unique for client ${client}, proceeding...`);

    setIsLoading(true);
    try {
      // Compute payload according to selection
      const couplingReplace = typeOfWO === 'Coupling' ? 'Yes' : 'No';
      const woType = typeOfWO === 'Coupling' ? 'Tubing' : pipeType;
      const payload: any = {
        wo_no: woNo,
        client: client,
        type: woType,
        diameter: typeOfWO === 'Coupling' ? '' : diameter,
        coupling_replace: couplingReplace,
        wo_date: woDate,
        transport: transport,
        key_col: keyCol,
        payer: '',
        planned_qty: typeOfWO === 'OCTG' ? plannedQty : '',
        price_type: typeOfWO === 'OCTG' ? priceType : '',
        price: typeOfWO === 'Coupling' ? price : (priceType === 'Fixed' ? price : ''),
        transportation_cost: transport === 'TCC' ? transportCost : '',
      };
      if (typeOfWO === 'OCTG' && priceType === 'Stage Based') {
        Object.assign(payload, stagePrices);
      }

      const success = await sharePointService.createWorkOrder(payload);
      
      if (success) {
        toast({
          title: "✅ Work Order создан успешно!",
          description: (
            <div className="space-y-2">
              <p className="font-bold text-white">
                Work Order <span className="bg-white text-green-600 px-2 py-1 rounded font-mono font-bold">{woNo}</span> для клиента <span className="bg-white text-blue-600 px-2 py-1 rounded font-bold">{client}</span>
              </p>
              <p className="text-sm text-white font-medium">
                🎉 Данные успешно сохранены в SharePoint Excel
              </p>
            </div>
          ),
          duration: 6000,
        });
        
        // Reset form but keep client and type-of-wo for faster entry
        const preservedClient = client;
        const preservedTypeOfWO = typeOfWO;
        setClient(preservedClient);
        setWoNo("");
        setWoDate("");
        setTypeOfWO(preservedTypeOfWO);
        setPipeType('Tubing');
        setDiameter("");
        setPlannedQty("");
        setPriceType('');
        setPrice("");
        setStagePrices({ rattling_price: "", external_price: "", hydro_price: "", mpi_price: "", drift_price: "", emi_price: "", marking_price: "" });
        setTransport('');
        setTransportCost("");

        // Refresh the work orders list to include the newly added WO (use preserved client)
        if (preservedClient) {
          console.log(`🔄 Refreshing work orders list for client ${preservedClient} after successful addition`);
          const updatedWorkOrders = await sharePointService.getWorkOrdersByClient(preservedClient);
          setExistingWorkOrders(updatedWorkOrders);
          console.log(`📋 Updated work orders list for ${preservedClient}:`, updatedWorkOrders);
        }

        // Auto-press "Update Data" button once: clear freshness and trigger background refresh
        try {
          if (sharePointService && refreshDataInBackground) {
            console.log('🟦 Auto Update Data: clearing last refresh and triggering background refresh');
            safeLocalStorage.removeItem('sharepoint_last_refresh');
            await refreshDataInBackground(sharePointService);
          }
        } catch (e) {
          console.warn('Auto Update Data encountered an error:', e);
        }
      } else {
        toast({
          title: "🔒 SharePoint файл заблокирован",
          description: (
            <div className="space-y-2">
              <p className="font-bold text-white">
                Кто-то редактирует Excel файл в данный момент
              </p>
              <p className="text-sm text-white font-medium">
                ⏳ Попробуйте через несколько минут, когда файл освободится
              </p>
            </div>
          ),
          variant: "destructive",
          duration: 8000,
        });
      }
    } catch (error) {
      console.error('Error creating work order:', error);
      toast({
        title: "Error",
        description: "Failed to create work order. Please try again.",
        variant: "destructive",
      });
    } finally {
      setIsLoading(false);
    }
  };

  // UI
  return (
    <div className="min-h-screen bg-gray-50">
      <Header />
      <div className="container mx-auto px-6 py-8">
        <div className="mb-6">
          <Button 
            variant="outline" 
            onClick={() => navigate("/")}
            className="flex items-center space-x-2 border-2 hover:bg-gray-50"
          >
            <ArrowLeft className="w-4 h-4" />
            <span>Back to Dashboard</span>
          </Button>
        </div>

        <Card className="max-w-4xl mx-auto border-2 shadow-lg">
          <CardHeader className="bg-blue-50 border-b-2">
            <CardTitle className="text-2xl font-bold text-blue-800">Add Work Order</CardTitle>
          </CardHeader>
          <CardContent className="p-6">
            <form onSubmit={handleSubmit} className="space-y-8">
              {/* Basic */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-2">
                  <Label className="text-sm font-semibold text-gray-700">Client</Label>
                  <Select value={client} onValueChange={(v) => setClient(v)}>
                    <SelectTrigger className="border-2 focus:border-blue-500 h-11">
                      <SelectValue placeholder="Select client" />
                    </SelectTrigger>
                    <SelectContent>
                      {availableClients.map(c => (
                        <SelectItem key={c} value={c}>{c}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-2">
                  <Label className="text-sm font-semibold text-gray-700">Work Order Number</Label>
                  <Input value={woNo} onChange={(e)=> setWoNo(e.target.value)} placeholder="Enter WO number" className="border-2 focus:border-blue-500 h-11" required />
                </div>
                <div className="space-y-2">
                  <Label className="text-sm font-semibold text-gray-700">WO Date</Label>
                  <DateInputField id="wo_date" value={woDate} onChange={(v)=> setWoDate(v)} className="border-2 focus:border-blue-500 h-11" placeholder="dd/mm/yyyy" />
                </div>
                <div className="space-y-2">
                  <Label className="text-sm font-semibold text-gray-700">Type of WO</Label>
                  <Select value={typeOfWO} onValueChange={(v)=> setTypeOfWO(v as any)}>
                    <SelectTrigger className="border-2 focus:border-blue-500 h-11">
                      <SelectValue placeholder="Select type of WO" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="OCTG">OCTG Inspection</SelectItem>
                      <SelectItem value="Coupling">Coupling Replace</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
              </div>

              {/* OCTG block */}
              {typeOfWO === 'OCTG' && (
                <div className="space-y-6">
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                    <div className="space-y-2">
                      <Label>Type Of Pipe</Label>
                      <Select value={pipeType} onValueChange={(v)=> setPipeType(v as any)}>
                        <SelectTrigger className="border-2 focus:border-blue-500 h-11"><SelectValue placeholder="Select pipe type" /></SelectTrigger>
                        <SelectContent>
                          <SelectItem value="Tubing">Tubing</SelectItem>
                          <SelectItem value="Sucker Rod">Sucker Rod</SelectItem>
                        </SelectContent>
                      </Select>
                    </div>
                    <div className="space-y-2">
                      <Label>Diameter</Label>
                      <Select value={diameter} onValueChange={(v)=> setDiameter(v)}>
                        <SelectTrigger className="border-2 focus:border-blue-500 h-11"><SelectValue placeholder="Select diameter" /></SelectTrigger>
                        <SelectContent>
                          <SelectItem value='3 1/2"'>3 1/2"</SelectItem>
                          <SelectItem value='2 7/8"'>2 7/8"</SelectItem>
                        </SelectContent>
                      </Select>
                    </div>
                    <div className="space-y-2">
                      <Label>Planned Qty</Label>
                      <Input inputMode="numeric" pattern="^[0-9]+$" value={plannedQty} onChange={(e)=> setPlannedQty(sanitizeInteger(e.target.value))} placeholder="Enter quantity" className="border-2 focus:border-blue-500 h-11" />
                    </div>
                  </div>

                  <div className="space-y-4">
                    <Label className="text-sm font-semibold text-gray-700">Price Type</Label>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div>
                        <Select value={priceType} onValueChange={(v)=> setPriceType(v as any)}>
                          <SelectTrigger className="border-2 focus:border-blue-500 h-11"><SelectValue placeholder="Select price type" /></SelectTrigger>
                          <SelectContent>
                            <SelectItem value="Fixed">Fixed</SelectItem>
                            <SelectItem value="Stage Based">Stage Based</SelectItem>
                          </SelectContent>
                        </Select>
                      </div>
                      {priceType === 'Fixed' && (
                        <div className="space-y-2">
                          <Label>Price for each pipe</Label>
                          <Input inputMode="decimal" pattern="^[0-9]*\.?[0-9]+$" value={price} onChange={(e)=> setPrice(sanitizeDecimal(e.target.value))} placeholder="e.g., 1250.50" className="border-2 focus:border-blue-500 h-11" />
                        </div>
                      )}
                    </div>

                    {priceType === 'Stage Based' && (
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        {[
                          { key:'rattling_price', label:'Item: 1.7 Rattling_Price' },
                          { key:'external_price', label:'Item: 1.1 External_Price' },
                          { key:'hydro_price', label:'Item: 1.2 Hydro_Price' },
                          { key:'mpi_price', label:'Item: 1.5 MPI_Price' },
                          { key:'drift_price', label:'Item: 1.3 Drift_Price' },
                          { key:'emi_price', label:'Item: 1.4 EMI_Price' },
                          { key:'marking_price', label:'Item: 1.6 Marking_Price' },
                        ].map((item)=> (
                          <div key={item.key} className="flex items-center justify-between gap-4 p-3 rounded border bg-yellow-50">
                            <span className="text-sm font-medium">{item.label}</span>
                            <Input inputMode="decimal" pattern="^[0-9]*\.?[0-9]+$" className="h-10 w-40"
                              value={(stagePrices as any)[item.key]}
                              onChange={(e)=> setStagePrices(prev => ({ ...prev, [item.key]: sanitizeDecimal(e.target.value) }))}
                              placeholder="0.00"
                            />
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              )}

              {/* Coupling Replace block */}
              {typeOfWO === 'Coupling' && (
                <div className="space-y-4">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div className="space-y-2">
                      <Label>Type Of Pipe</Label>
                      <Input value="Tubing" readOnly className="h-11 w-full rounded-md border border-gray-300 bg-gray-100 px-3 text-gray-500 shadow-sm cursor-not-allowed" />
                    </div>
                    <div className="space-y-2">
                      <Label>Price for replacement</Label>
                      <Input inputMode="decimal" pattern="^[0-9]*\.?[0-9]+$" value={price} onChange={(e)=> setPrice(sanitizeDecimal(e.target.value))} placeholder="e.g., 7000.00" className="border-2 focus:border-blue-500 h-11" />
                    </div>
                  </div>
                </div>
              )}

              {/* Transport */}
              <div className="space-y-2">
                <Label className="text-sm font-semibold text-gray-700">Transport</Label>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div>
                    <Select value={transport} onValueChange={(v)=> setTransport(v as any)}>
                      <SelectTrigger className="border-2 focus:border-blue-500 h-11"><SelectValue placeholder="Select transport" /></SelectTrigger>
                      <SelectContent>
                        <SelectItem value="Client">Client</SelectItem>
                        <SelectItem value="TCC">TCC</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>
                  {transport === 'TCC' && (
                    <div className="space-y-2">
                      <Label>Transportation Cost</Label>
                      <Input inputMode="decimal" pattern="^[0-9]*\.?[0-9]+$" value={transportCost} onChange={(e)=> setTransportCost(sanitizeDecimal(e.target.value))} placeholder="0.00" className="border-2 focus:border-blue-500 h-11" />
                    </div>
                  )}
                </div>
              </div>

              {/* Key */}
              <div className="space-y-2">
                <Label className="text-sm font-semibold text-gray-700">Key Column</Label>
                <Input value={keyCol} placeholder="Auto-generated based on WO, Client, Type, Diameter" className="h-11 w-full rounded-md border border-gray-300 bg-gray-100 px-3 text-gray-500 shadow-sm cursor-not-allowed" readOnly />
              </div>

              <div className="flex justify-end space-x-4 pt-6 border-t-2 border-gray-100">
                <Button type="button" variant="outline" onClick={() => navigate("/")} className="border-2 h-12 px-6">Cancel</Button>
                <Button type="submit" className="h-12 px-6 font-semibold" disabled={isLoading || !isConnected}>
                  {isLoading ? "Creating..." : "Save Work Order"}
                </Button>
              </div>
            </form>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
