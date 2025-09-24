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
import { useSharePointInstantData } from "@/hooks/useInstantData";
import { useToast } from "@/hooks/use-toast";
import { DateInputField } from "@/components/ui/date-input";
import { safeLocalStorage } from '@/lib/safe-storage';

export default function TubingForm() {
  const navigate = useNavigate();
  const [formData, setFormData] = useState({
    client: "",
    wo_no: "",
    batch: "",
    diameter: "",
    qty: "",
    pipe_from: "",
    pipe_to: "",
    class_1: "",
    class_2: "",
    class_3: "",
    repair: "",
    scrap: "",
    start_date: "",
    end_date: "",
    rattling_qty: "",
    external_qty: "",
    hydro_qty: "",
    mpi_qty: "",
    drift_qty: "",
    emi_qty: "",
    marking_qty: "",
    act_no_oper: "",
    act_date: ""
  });

  const [availableClients, setAvailableClients] = useState<string[]>([]);
  const [availableWorkOrders, setAvailableWorkOrders] = useState<string[]>([]);
  const [availableDiameters, setAvailableDiameters] = useState<string[]>([]);
  const [nextBatch, setNextBatch] = useState<string>("");
  const [lastPipeTo, setLastPipeTo] = useState<number>(0);
  const [isLoading, setIsLoading] = useState(false);
  const { user } = useAuth();
  const { sharePointService, isConnected } = useSharePoint();
  const { clients } = useSharePointInstantData();
  const { toast } = useToast();

  // Мгновенная загрузка из кеша - всегда используем кешированные данные
  useEffect(() => {
    const filteredClients = clients.filter(client => client && client.trim());
    if (filteredClients.length > 0) {
      setAvailableClients(filteredClients);
      console.log('⚡ TubingForm loaded clients from cache:', filteredClients.length);
    }
  }, [clients]);

  // Load existing work orders when client is selected - same logic as WOForm
  useEffect(() => {
    const loadWorkOrders = async () => {
      if (!formData.client || !sharePointService) {
        setAvailableWorkOrders([]);
        return;
      }
      
      try {
        console.log(`🔍 Loading work orders for client: ${formData.client}`);
        const workOrders = await sharePointService.getWorkOrdersByClient(formData.client);
        console.log(`📋 Found ${workOrders.length} existing work orders for ${formData.client}:`, workOrders);
        setAvailableWorkOrders(workOrders);
      } catch (error) {
        console.error('❌ Error loading work orders:', error);
        setAvailableWorkOrders([]);
      }
    };
    loadWorkOrders();
  }, [formData.client, sharePointService]);

  // INSTANT batch calculation - no delays, no glitches
  useEffect(() => {
    if (!formData.client || !formData.wo_no) {
      setNextBatch("");
      setAvailableDiameters([]);
      setLastPipeTo(0);
      return;
    }
    
    // INSTANT calculation from cache only - no API calls
    const cachedTubingData = safeLocalStorage.getItem('sharepoint_cached_tubing');
    const cachedWoData = safeLocalStorage.getItem('sharepoint_cached_workorders');
    
    if (cachedTubingData && cachedWoData) {
      try {
        const tubingData = JSON.parse(cachedTubingData);
        const woData = JSON.parse(cachedWoData);
        
        // INSTANT batch calculation
        if (!tubingData || tubingData.length === 0) {
          setNextBatch("Batch # 1");
          setLastPipeTo(0);
        } else {
          const headers = tubingData[0];
          const clientIndex = headers.findIndex((h: string) => h && h.toLowerCase().includes('client'));
          const woIndex = headers.findIndex((h: string) => h && h.toLowerCase().includes('wo'));
          const batchIndex = headers.findIndex((h: string) => h && h.toLowerCase().includes('batch'));
          const pipeToIndex = headers.findIndex((h: string) => h && h.toLowerCase().includes('pipe_to'));

          const clientWoRecords = tubingData.slice(1).filter(row => 
            row[clientIndex] === formData.client && row[woIndex] === formData.wo_no
          );

          let maxBatchNum = 0;
          let maxPipeTo = 0;

          clientWoRecords.forEach(row => {
            const batchStr = row[batchIndex];
            if (batchStr && typeof batchStr === 'string') {
              const batchMatch = batchStr.match(/(\d+)/);
              if (batchMatch) {
                const batchNum = parseInt(batchMatch[1]);
                if (batchNum > maxBatchNum) {
                  maxBatchNum = batchNum;
                  maxPipeTo = parseInt(row[pipeToIndex]) || 0;
                }
              }
            }
          });

          const nextBatchNum = maxBatchNum + 1;
          setNextBatch(`Batch # ${nextBatchNum}`);
          setLastPipeTo(maxPipeTo);
          console.log(`📊 Next batch: ${nextBatchNum}, Last Pipe_To: ${maxPipeTo}`);
        }

        // Get diameter from work order data
        if (woData && woData.length > 0) {
          const woHeaders = woData[0];
          const woClientIndex = woHeaders.findIndex((h: string) => h && h.toLowerCase().includes('client'));
          const woNoIndex = woHeaders.findIndex((h: string) => h && h.toLowerCase().includes('wo'));
          const diameterIndex = woHeaders.findIndex((h: string) => h && h.toLowerCase().includes('diameter'));

          const woRecord = woData.slice(1).find(row => 
            row[woClientIndex] === formData.client && row[woNoIndex] === formData.wo_no
          );

          if (woRecord && woRecord[diameterIndex]) {
            setAvailableDiameters([woRecord[diameterIndex]]);
            setFormData(prev => ({ ...prev, diameter: woRecord[diameterIndex] }));
          }
        }

      } catch (error) {
        console.error('❌ Error loading batch info:', error);
        setNextBatch("Batch # 1");
        setLastPipeTo(0);
      }
    }
  }, [formData.client, formData.wo_no]);

  // Auto-calculate Pipe_From and Pipe_To when quantity changes
  useEffect(() => {
    if (formData.qty && !isNaN(parseInt(formData.qty))) {
      const qty = parseInt(formData.qty);
      const pipeFrom = lastPipeTo + 1;
      const pipeTo = lastPipeTo + qty;
      
      setFormData(prev => ({
        ...prev,
        pipe_from: pipeFrom.toString(),
        pipe_to: pipeTo.toString()
      }));
    }
  }, [formData.qty, lastPipeTo]);

  // Set batch when it's calculated
  useEffect(() => {
    if (nextBatch) {
      setFormData(prev => ({ ...prev, batch: nextBatch }));
    }
  }, [nextBatch]);

  const validateInspectionQuantities = (): boolean => {
    if (!formData.qty) return true;
    
    const totalQty = parseInt(formData.qty);
    const inspectionFields = [
      { name: 'Rattling Qty', value: formData.rattling_qty },
      { name: 'External Qty', value: formData.external_qty },
      { name: 'Hydro Qty', value: formData.hydro_qty },
      { name: 'MPI Qty', value: formData.mpi_qty },
      { name: 'Drift Qty', value: formData.drift_qty },
      { name: 'EMI Qty', value: formData.emi_qty },
      { name: 'Marking Qty', value: formData.marking_qty }
    ];

    for (const field of inspectionFields) {
      if (field.value && !isNaN(parseInt(field.value))) {
        const qty = parseInt(field.value);
        if (qty > totalQty) {
          toast({
            title: "❌ Ошибка валидации",
            description: `${field.name} (${qty}) не может быть больше общего количества труб (${totalQty})`,
            variant: "destructive",
          });
          return false;
        }
      }
    }
    return true;
  };

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
    if (!formData.client || !formData.wo_no || !formData.qty) {
      toast({
        title: "Ошибка валидации",
        description: "Заполните клиента, Work Order и количество",
        variant: "destructive",
      });
      return;
    }

    // Validate inspection quantities
    if (!validateInspectionQuantities()) {
      return;
    }

    setIsLoading(true);
    try {
      const success = await sharePointService.addTubingRecordToExcel({
        client: formData.client,
        wo_no: formData.wo_no,
        batch: formData.batch,
        diameter: formData.diameter,
        qty: formData.qty,
        pipe_from: formData.pipe_from,
        pipe_to: formData.pipe_to,
        class_1: formData.class_1,
        class_2: formData.class_2,
        class_3: formData.class_3,
        repair: formData.repair,
        scrap: formData.scrap,
        start_date: formData.start_date,
        end_date: formData.end_date,
        rattling_qty: formData.rattling_qty || "0",
        external_qty: formData.external_qty || "0",
        hydro_qty: formData.hydro_qty || "0",
        mpi_qty: formData.mpi_qty || "0",
        drift_qty: formData.drift_qty || "0",
        emi_qty: formData.emi_qty || "0",
        marking_qty: formData.marking_qty || "0",
        act_no_oper: formData.act_no_oper,
        act_date: formData.act_date
      });
      
      if (success) {
        toast({
          title: "✅ Tubing record создан успешно!",
          description: (
            <div className="space-y-2">
              <p className="font-bold text-white">
                {formData.batch} для клиента <span className="bg-white text-blue-600 px-2 py-1 rounded font-bold">{formData.client}</span>
              </p>
              <p className="text-sm text-white font-medium">
                🎉 Данные успешно сохранены в SharePoint Excel
              </p>
            </div>
          ),
          duration: 6000,
        });
        
        // Reset form
        setFormData({
          client: "",
          wo_no: "",
          batch: "",
          diameter: "",
          qty: "",
          pipe_from: "",
          pipe_to: "",
          class_1: "",
          class_2: "",
          class_3: "",
          repair: "",
          scrap: "",
          start_date: "",
          end_date: "",
          rattling_qty: "",
          external_qty: "",
          hydro_qty: "",
          mpi_qty: "",
          drift_qty: "",
          emi_qty: "",
          marking_qty: "",
          act_no_oper: "",
          act_date: ""
        });
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
      console.error('Error creating tubing record:', error);
      toast({
        title: "Error",
        description: "Failed to create tubing record. Please try again.",
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
          <CardHeader className="bg-green-50 border-b-2">
            <CardTitle className="text-2xl font-bold text-green-800">Tubing Registry</CardTitle>
          </CardHeader>
          <CardContent>
            <form onSubmit={handleSubmit} className="space-y-6">
              {/* Basic Information */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="client">Client *</Label>
                  <Select value={formData.client} onValueChange={(value) => handleInputChange("client", value)}>
                    <SelectTrigger>
                      <SelectValue placeholder="Select client" />
                    </SelectTrigger>
                    <SelectContent>
                      {availableClients.map((client) => (
                        <SelectItem key={client} value={client}>
                          {client}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                <div>
                  <Label htmlFor="wo_no">Work Order Number *</Label>
                  <Select 
                    value={formData.wo_no || ""} 
                    onValueChange={(value) => handleInputChange("wo_no", value)}
                    disabled={!formData.client}
                  >
                    <SelectTrigger className="border-2 focus:border-blue-500 h-11">
                      <SelectValue placeholder={formData.wo_no || "Select work order"} />
                    </SelectTrigger>
                    <SelectContent>
                      {availableWorkOrders.map((wo) => (
                        <SelectItem key={wo} value={wo}>
                          {wo}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                <div>
                  <Label htmlFor="batch">Batch Number</Label>
                  <Input
                    id="batch"
                    value={nextBatch || formData.batch}
                    readOnly
                    className="h-11 w-full rounded-md border border-gray-300 bg-gray-100 px-3 text-gray-500 shadow-sm cursor-not-allowed"
                    placeholder="Auto-calculated"
                  />
                </div>

                <div>
                  <Label htmlFor="diameter">Diameter *</Label>
                  <Select 
                    value={formData.diameter} 
                    onValueChange={(value) => handleInputChange("diameter", value)}
                    disabled={!formData.wo_no}
                  >
                    <SelectTrigger>
                      <SelectValue placeholder="Select diameter" />
                    </SelectTrigger>
                    <SelectContent>
                      {availableDiameters.map((diameter) => (
                        <SelectItem key={diameter} value={diameter}>
                          {diameter}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                <div>
                  <Label htmlFor="qty">Quantity *</Label>
                  <Input
                    id="qty"
                    type="number"
                    value={formData.qty}
                    onChange={(e) => handleInputChange("qty", e.target.value)}
                    placeholder="Enter quantity"
                    required
                  />
                </div>

                <div>
                  <Label htmlFor="pipe_from">Pipe From</Label>
                  <Input
                    id="pipe_from"
                    value={formData.pipe_from}
                    onChange={(e) => handleInputChange("pipe_from", e.target.value)}
                    placeholder="Auto-calculated"
                    readOnly
                    className="h-11 w-full rounded-md border border-gray-300 bg-gray-100 px-3 text-gray-500 shadow-sm cursor-not-allowed"
                  />
                </div>

                <div>
                  <Label htmlFor="pipe_to">Pipe To</Label>
                  <Input
                    id="pipe_to"
                    value={formData.pipe_to}
                    onChange={(e) => handleInputChange("pipe_to", e.target.value)}
                    placeholder="Auto-calculated"
                    readOnly
                    className="h-11 w-full rounded-md border border-gray-300 bg-gray-100 px-3 text-gray-500 shadow-sm cursor-not-allowed"
                  />
                </div>
              </div>

              {/* Classification */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <Label htmlFor="class_1">Class 1</Label>
                  <Input
                    id="class_1"
                    value={formData.class_1}
                    onChange={(e) => handleInputChange("class_1", e.target.value)}
                    placeholder="Class 1"
                  />
                </div>

                <div>
                  <Label htmlFor="class_2">Class 2</Label>
                  <Input
                    id="class_2"
                    value={formData.class_2}
                    onChange={(e) => handleInputChange("class_2", e.target.value)}
                    placeholder="Class 2"
                  />
                </div>

                <div>
                  <Label htmlFor="class_3">Class 3</Label>
                  <Input
                    id="class_3"
                    value={formData.class_3}
                    onChange={(e) => handleInputChange("class_3", e.target.value)}
                    placeholder="Class 3"
                  />
                </div>
              </div>

              {/* Repair and Scrap */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="repair">Repair</Label>
                  <Input
                    id="repair"
                    value={formData.repair}
                    onChange={(e) => handleInputChange("repair", e.target.value)}
                    placeholder="Repair"
                  />
                </div>

                <div>
                  <Label htmlFor="scrap">Scrap</Label>
                  <Input
                    id="scrap"
                    value={formData.scrap}
                    onChange={(e) => handleInputChange("scrap", e.target.value)}
                    placeholder="Scrap"
                  />
                </div>
              </div>

              {/* Dates */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="start_date">Start Date</Label>
                  <DateInputField
                    id="start_date"
                    value={formData.start_date}
                    onChange={(value) => handleInputChange("start_date", value)}
                    placeholder="dd/mm/yyyy"
                  />
                </div>

                <div>
                  <Label htmlFor="end_date">End Date</Label>
                  <DateInputField
                    id="end_date"
                    value={formData.end_date}
                    onChange={(value) => handleInputChange("end_date", value)}
                    placeholder="dd/mm/yyyy"
                  />
                </div>
              </div>

              {/* Inspection Quantities */}
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                <div>
                  <Label htmlFor="rattling_qty">Rattling Qty</Label>
                  <Input
                    id="rattling_qty"
                    type="number"
                    value={formData.rattling_qty}
                    onChange={(e) => handleInputChange("rattling_qty", e.target.value)}
                    placeholder="0"
                  />
                </div>

                <div>
                  <Label htmlFor="external_qty">External Qty</Label>
                  <Input
                    id="external_qty"
                    type="number"
                    value={formData.external_qty}
                    onChange={(e) => handleInputChange("external_qty", e.target.value)}
                    placeholder="0"
                  />
                </div>

                <div>
                  <Label htmlFor="hydro_qty">Hydro Qty</Label>
                  <Input
                    id="hydro_qty"
                    type="number"
                    value={formData.hydro_qty}
                    onChange={(e) => handleInputChange("hydro_qty", e.target.value)}
                    placeholder="0"
                  />
                </div>

                <div>
                  <Label htmlFor="mpi_qty">MPI Qty</Label>
                  <Input
                    id="mpi_qty"
                    type="number"
                    value={formData.mpi_qty}
                    onChange={(e) => handleInputChange("mpi_qty", e.target.value)}
                    placeholder="0"
                  />
                </div>

                <div>
                  <Label htmlFor="drift_qty">Drift Qty</Label>
                  <Input
                    id="drift_qty"
                    type="number"
                    value={formData.drift_qty}
                    onChange={(e) => handleInputChange("drift_qty", e.target.value)}
                    placeholder="0"
                  />
                </div>

                <div>
                  <Label htmlFor="emi_qty">EMI Qty</Label>
                  <Input
                    id="emi_qty"
                    type="number"
                    value={formData.emi_qty}
                    onChange={(e) => handleInputChange("emi_qty", e.target.value)}
                    placeholder="0"
                  />
                </div>

                <div>
                  <Label htmlFor="marking_qty">Marking Qty</Label>
                  <Input
                    id="marking_qty"
                    type="number"
                    value={formData.marking_qty}
                    onChange={(e) => handleInputChange("marking_qty", e.target.value)}
                    placeholder="0"
                  />
                </div>
              </div>

              {/* Activity Information */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="act_no_oper">Activity Number/Operator</Label>
                  <Input
                    id="act_no_oper"
                    value={formData.act_no_oper}
                    onChange={(e) => handleInputChange("act_no_oper", e.target.value)}
                    placeholder="Activity number or operator"
                  />
                </div>

                <div>
                  <Label htmlFor="act_date">Activity Date</Label>
                  <DateInputField
                    id="act_date"
                    value={formData.act_date}
                    onChange={(value) => handleInputChange("act_date", value)}
                    placeholder="dd/mm/yyyy"
                  />
                </div>
              </div>

              <div className="flex justify-end space-x-4 pt-6 border-t-2 border-gray-100">
                <Button type="button" variant="outline" onClick={() => navigate("/")} className="border-2 h-12 px-6">
                  Cancel
                </Button>
                <Button type="submit" className="h-12 px-6 font-semibold" disabled={isLoading || !isConnected}>
                  {isLoading ? "Creating..." : "Add Tubing Record"}
                </Button>
              </div>
            </form>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
