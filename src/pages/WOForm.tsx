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
  const [formData, setFormData] = useState({
    wo_no: "",
    client: "",
    type: "",
    diameter: "",
    coupling_replace: "",
    wo_date: "",
    transport: "",
    key_col: "",
    payer: "",
    planned_qty: ""
  });

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
      if (!formData.client || !sharePointService) {
        setExistingWorkOrders([]);
        return;
      }
      
      try {
        console.log(`🔍 Loading work orders for client: ${formData.client}`);
        const workOrders = await sharePointService.getWorkOrdersByClient(formData.client);
        console.log(`📋 Found ${workOrders.length} existing work orders for ${formData.client}:`, workOrders);
        setExistingWorkOrders(workOrders);
      } catch (error) {
        console.error('❌ Error loading work orders:', error);
        setExistingWorkOrders([]);
      }
    };
    loadWorkOrders();
  }, [formData.client, sharePointService]);

  // Auto-generate Key Col when relevant fields change
  useEffect(() => {
    if (formData.wo_no && formData.client && formData.type && formData.diameter) {
      const keyCol = `${formData.wo_no} - ${formData.client} - ${formData.type} - ${formData.diameter}`;
      setFormData(prev => ({ ...prev, key_col: keyCol }));
    }
  }, [formData.wo_no, formData.client, formData.type, formData.diameter]);

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
    if (!formData.wo_no || !formData.client) {
      toast({
        title: "Ошибка валидации",
        description: "Заполните номер Work Order и выберите клиента",
        variant: "destructive",
      });
      return;
    }

    // Check if work order already exists for this client
    console.log(`🔍 Checking duplicate: WO ${formData.wo_no} for client ${formData.client}`);
    console.log(`📋 Existing work orders for ${formData.client}:`, existingWorkOrders);
    
    const isDuplicate = existingWorkOrders.some(wo => wo.toString().trim() === formData.wo_no.toString().trim());
    console.log(`🔍 Duplicate check result:`, isDuplicate);
    
    if (isDuplicate) {
      console.log(`❌ DUPLICATE FOUND: Work Order ${formData.wo_no} already exists for client ${formData.client}`);
        toast({
          title: "🚫 Work Order уже существует",
          description: (
            <div className="space-y-2">
              <p className="font-bold text-white">
                Work Order <span className="bg-white text-red-600 px-2 py-1 rounded font-mono font-bold">{formData.wo_no}</span> уже существует для клиента <span className="bg-white text-blue-600 px-2 py-1 rounded font-bold">{formData.client}</span>
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
    
    console.log(`✅ Work Order ${formData.wo_no} is unique for client ${formData.client}, proceeding...`);

    setIsLoading(true);
    try {
      const success = await sharePointService.createWorkOrder({
        wo_no: formData.wo_no,
        client: formData.client,
        type: formData.type,
        diameter: formData.diameter,
        coupling_replace: formData.coupling_replace,
        wo_date: formData.wo_date,
        transport: formData.transport,
        key_col: formData.key_col,
        payer: formData.payer,
        planned_qty: formData.planned_qty
      });
      
      if (success) {
        toast({
          title: "✅ Work Order создан успешно!",
          description: (
            <div className="space-y-2">
              <p className="font-bold text-white">
                Work Order <span className="bg-white text-green-600 px-2 py-1 rounded font-mono font-bold">{formData.wo_no}</span> для клиента <span className="bg-white text-blue-600 px-2 py-1 rounded font-bold">{formData.client}</span>
              </p>
              <p className="text-sm text-white font-medium">
                🎉 Данные успешно сохранены в SharePoint Excel
              </p>
            </div>
          ),
          duration: 6000,
        });
        
        // Reset form but keep frequently reused fields to allow adding next WO without reload
        const preservedClient = formData.client;
        const preservedType = formData.type;
        const preservedDiameter = formData.diameter;
        const preservedCoupling = formData.coupling_replace;
        setFormData({
          wo_no: '',
          client: preservedClient,
          type: preservedType,
          diameter: preservedDiameter,
          coupling_replace: preservedCoupling,
          wo_date: '',
          transport: '',
          key_col: '',
          payer: '',
          planned_qty: ''
        });

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

  const handleInputChange = (field: string, value: string) => {
    setFormData(prev => ({ ...prev, [field]: value }));
  };

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
            <form onSubmit={handleSubmit} className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-2">
                  <Label htmlFor="wo_no" className="text-sm font-semibold text-gray-700">Work Order Number</Label>
                  <Input
                    id="wo_no"
                    value={formData.wo_no}
                    onChange={(e) => handleInputChange("wo_no", e.target.value)}
                    placeholder="Enter WO number"
                    className="border-2 focus:border-blue-500 h-11"
                    required
                  />
                </div>

                <div className="space-y-2">
                  <Label htmlFor="client" className="text-sm font-semibold text-gray-700">Client</Label>
                  <Select value={formData.client} onValueChange={(value) => handleInputChange("client", value)}>
                    <SelectTrigger className="border-2 focus:border-blue-500 h-11">
                      <SelectValue placeholder="Select client" />
                    </SelectTrigger>
                    <SelectContent>
                      {availableClients.map(client => (
                        <SelectItem key={client} value={client}>{client}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                  <p className="text-xs text-blue-600 font-medium">
                    Чтобы добавить новый Client, создайте запись в базе данных
                  </p>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="type" className="text-sm font-semibold text-gray-700">Type</Label>
                  <Select onValueChange={(value) => handleInputChange("type", value)} value={formData.type}>
                    <SelectTrigger className="border-2 focus:border-blue-500 h-11">
                      <SelectValue placeholder="Select type" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="Tubing">Tubing</SelectItem>
                      <SelectItem value="Sucker Rod">Sucker Rod</SelectItem>
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="diameter" className="text-sm font-semibold text-gray-700">Diameter</Label>
                  <Select value={formData.diameter} onValueChange={(value) => handleInputChange("diameter", value)}>
                    <SelectTrigger className="border-2 focus:border-blue-500 h-11">
                      <SelectValue placeholder="Select diameter" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="3 1/2&quot;">3 1/2"</SelectItem>
                      <SelectItem value="2 7/8&quot;">2 7/8"</SelectItem>
                    </SelectContent>
                  </Select>
                  <p className="text-xs text-blue-600 font-medium">
                    Стандартные диаметры для трубной продукции
                  </p>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="coupling_replace" className="text-sm font-semibold text-gray-700">Coupling Replace</Label>
                  <Select onValueChange={(value) => handleInputChange("coupling_replace", value)} value={formData.coupling_replace}>
                    <SelectTrigger className="border-2 focus:border-blue-500 h-11">
                      <SelectValue placeholder="Select option" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="No">No</SelectItem>
                      <SelectItem value="Yes">Yes</SelectItem>
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="wo_date" className="text-sm font-semibold text-gray-700">Work Order Date</Label>
                  <DateInputField
                    id="wo_date"
                    value={formData.wo_date}
                    onChange={(v) => handleInputChange("wo_date", v)}
                    className="border-2 focus:border-blue-500 h-11"
                    placeholder="dd/mm/yyyy"
                  />
                </div>

                <div className="space-y-2">
                  <Label htmlFor="transport" className="text-sm font-semibold text-gray-700">Transport</Label>
                  <Input
                    id="transport"
                    value={formData.transport}
                    onChange={(e) => handleInputChange("transport", e.target.value)}
                    placeholder="Enter transport details"
                    className="border-2 focus:border-blue-500 h-11"
                  />
                </div>

                <div className="space-y-2">
                  <Label htmlFor="key_col" className="text-sm font-semibold text-gray-700">Key Column</Label>
                  <Input
                    id="key_col"
                    value={formData.key_col}
                    placeholder="Auto-generated based on WO, Client, Type, Diameter"
                    className="h-11 w-full rounded-md border border-gray-300 bg-gray-100 px-3 text-gray-500 shadow-sm cursor-not-allowed"
                    readOnly
                  />
                  <p className="text-xs text-blue-600 font-medium">
                    Автоматически генерируется: WO - Client - Type - Diameter
                  </p>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="payer" className="text-sm font-semibold text-gray-700">Payer</Label>
                  <Input
                    id="payer"
                    value={formData.payer}
                    onChange={(e) => handleInputChange("payer", e.target.value)}
                    placeholder="Enter payer information"
                    className="border-2 focus:border-blue-500 h-11"
                  />
                </div>

                <div className="space-y-2">
                  <Label htmlFor="planned_qty" className="text-sm font-semibold text-gray-700">Planned Quantity</Label>
                  <Input
                    id="planned_qty"
                    type="number"
                    value={formData.planned_qty}
                    onChange={(e) => handleInputChange("planned_qty", e.target.value)}
                    placeholder="Enter planned quantity"
                    className="border-2 focus:border-blue-500 h-11"
                    required
                  />
                </div>
              </div>

              <div className="flex justify-end space-x-4 pt-6 border-t-2 border-gray-100">
                <Button type="button" variant="outline" onClick={() => navigate("/")} className="border-2 h-12 px-6">
                  Cancel
                </Button>
                <Button type="submit" className="h-12 px-6 font-semibold" disabled={isLoading || !isConnected}>
                  {isLoading ? "Creating..." : "Create Work Order"}
                </Button>
              </div>
            </form>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
