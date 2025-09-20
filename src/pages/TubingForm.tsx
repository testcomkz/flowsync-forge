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
  const { clients, workOrders: cachedWorkOrders, tubingData: cachedTubingData } = useSharePointInstantData();
  const { toast } = useToast();

  // Мгновенная загрузка из кеша - всегда используем кешированные данные
  useEffect(() => {
    const filteredClients = clients.filter(client => client && client.trim());
    if (filteredClients.length > 0) {
      setAvailableClients(filteredClients);
      console.log('⚡ TubingForm loaded clients from cache:', filteredClients.length);
    }
  }, [clients]);

  // Load existing work orders when client is selected - use cached data first
  useEffect(() => {
    if (!formData.client) {
      setAvailableWorkOrders([]);
      return;
    }

    console.log(`🔍 Loading work orders for client: ${formData.client}`);
    
    // Use cached work orders data first
    if (cachedWorkOrders && cachedWorkOrders.length > 0) {
      try {
        const headers = cachedWorkOrders[0];
        const clientIndex = headers.findIndex((h: string) => h && h.toLowerCase().includes('client'));
        const woIndex = headers.findIndex((h: string) => h && h.toLowerCase().includes('wo'));
        
        if (clientIndex !== -1 && woIndex !== -1) {
          let clientWorkOrders = Array.from(new Set(
            cachedWorkOrders
              .slice(1)
              .filter(row => String(row[clientIndex]).trim() === String(formData.client).trim())
              .map(row => String(row[woIndex]).trim())
              .filter(wo => !!wo)
          ));
          // Сохраняем выбранный WO в списке (Radix Select требует наличие item для отображения value)
          if (formData.wo_no && !clientWorkOrders.includes(String(formData.wo_no).trim())) {
            clientWorkOrders = [String(formData.wo_no).trim(), ...clientWorkOrders];
          }
          
          console.log(`📋 Found ${clientWorkOrders.length} work orders for ${formData.client}:`, clientWorkOrders);
          setAvailableWorkOrders(clientWorkOrders);
          return;
        }
      } catch (error) {
        console.error('❌ Error processing cached work orders:', error);
      }
    }

    // Fallback to API call if cache is not available
    if (sharePointService) {
      sharePointService.getWorkOrdersByClient(formData.client)
        .then(workOrders => {
          let normalized = Array.from(new Set(
            (workOrders || []).map(wo => String(wo).trim()).filter(Boolean)
          ));
          if (formData.wo_no && !normalized.includes(String(formData.wo_no).trim())) {
            normalized = [String(formData.wo_no).trim(), ...normalized];
          }
          console.log(`📋 API: Found ${normalized.length} work orders for ${formData.client}:`, normalized);
          setAvailableWorkOrders(normalized);
        })
        .catch(error => {
          console.error('❌ Error loading work orders from API:', error);
          setAvailableWorkOrders([]);
        });
    }
  }, [formData.client, cachedWorkOrders, sharePointService]);

  // INSTANT batch calculation with cached data
  useEffect(() => {
    console.log('🔄 Batch calculation triggered for:', { client: formData.client, wo_no: formData.wo_no });
    
    if (!formData.client || !formData.wo_no) {
      console.log('❌ Missing client or WO, resetting batch');
      setNextBatch("");
      setAvailableDiameters([]);
      setLastPipeTo(0);
      return;
    }
    
    // Use cached data from useSharePointInstantData hook
    if (cachedTubingData && cachedTubingData.length > 0 && cachedWorkOrders && cachedWorkOrders.length > 0) {
      try {
        console.log('📋 Using cached data for batch calculation');
        
        // INSTANT batch calculation (lowest missing positive integer)
        const tubingHeaders = cachedTubingData[0];
        const clientIndex = tubingHeaders.findIndex((h: string) => h && h.toLowerCase().includes('client'));
        const woIndex = tubingHeaders.findIndex((h: string) => h && h.toLowerCase().includes('wo'));
        const batchIndex = tubingHeaders.findIndex((h: string) => h && h.toLowerCase().includes('batch'));
        const pipeToIndex = tubingHeaders.findIndex((h: string) => h && h.toLowerCase().includes('pipe_to'));

        console.log('📍 Tubing column indexes:', { clientIndex, woIndex, batchIndex, pipeToIndex });

        const clientWoRecords = cachedTubingData.slice(1).filter(row => 
          String(row[clientIndex]).trim() === String(formData.client).trim() &&
          String(row[woIndex]).trim() === String(formData.wo_no).trim()
        );

        console.log(`🔍 Found ${clientWoRecords.length} existing records for ${formData.client}, WO ${formData.wo_no}`);

        const numsSet = new Set<number>();
        const pipeToByBatch = new Map<number, number>();
        clientWoRecords.forEach(row => {
          const raw = row[batchIndex];
          const num = raw != null ? parseInt(String(raw).match(/(\d+)/)?.[1] || '') : NaN;
          if (!isNaN(num) && num > 0) {
            numsSet.add(num);
            if (pipeToIndex !== -1) {
              const pt = parseInt(String(row[pipeToIndex]).replace(/[^\d-]/g, ''));
              pipeToByBatch.set(num, isNaN(pt) ? 0 : pt);
            }
          }
        });

        // Next batch = (max existing) + 1, не заполняем пропуски
        const nums = Array.from(numsSet);
        const maxNum = nums.length ? Math.max(...nums) : 0;
        const nextNum = maxNum + 1;
        const lastFromMax = maxNum > 0 ? (pipeToByBatch.get(maxNum) || 0) : 0;

        const calculatedBatch = `Batch # ${nextNum}`;
        console.log(`🎯 FINAL BATCH: ${calculatedBatch}, lastPipeTo from batch ${maxNum}: ${lastFromMax}`);
        setNextBatch(calculatedBatch);
        setLastPipeTo(lastFromMax);

        // Get diameter from work order data
        const woHeaders = cachedWorkOrders[0];
        const woClientIndex = woHeaders.findIndex((h: string) => h && h.toLowerCase().includes('client'));
        const woNoIndex = woHeaders.findIndex((h: string) => h && h.toLowerCase().includes('wo'));
        const diameterIndex = woHeaders.findIndex((h: string) => h && h.toLowerCase().includes('diameter'));

        const woRecord = cachedWorkOrders.slice(1).find(row => 
          String(row[woClientIndex]).trim() === String(formData.client).trim() &&
          String(row[woNoIndex]).trim() === String(formData.wo_no).trim()
        );

        if (woRecord && woRecord[diameterIndex]) {
          console.log('📏 Setting diameter:', woRecord[diameterIndex]);
          setAvailableDiameters([woRecord[diameterIndex]]);
          setFormData(prev => ({ ...prev, diameter: woRecord[diameterIndex] }));
        }

      } catch (error) {
        console.error('❌ Error in batch calculation:', error);
        setNextBatch("Batch # 1");
        setLastPipeTo(0);
      }
    } else {
      console.log('❌ No cached data available, setting default batch');
      setNextBatch("Batch # 1");
      setLastPipeTo(0);
    }
  }, [formData.client, formData.wo_no, cachedTubingData, cachedWorkOrders]);

  // Helper: recompute next batch from cache to lock on submit
  const recomputeNextBatchInfo = () => {
    if (!formData.client || !formData.wo_no || !cachedTubingData || cachedTubingData.length === 0) {
      return { nextBatchLabel: 'Batch # 1', nextBatchNum: 1, lastPipeTo: 0 };
    }
    try {
      const tubingHeaders = cachedTubingData[0];
      const clientIndex = tubingHeaders.findIndex((h: string) => h && h.toLowerCase().includes('client'));
      const woIndex = tubingHeaders.findIndex((h: string) => h && h.toLowerCase().includes('wo'));
      const batchIndex = tubingHeaders.findIndex((h: string) => h && h.toLowerCase().includes('batch'));
      const pipeToIndex = tubingHeaders.findIndex((h: string) => h && h.toLowerCase().includes('pipe_to'));

      const clientWoRecords = cachedTubingData.slice(1).filter(row => 
        String(row[clientIndex]).trim() === String(formData.client).trim() &&
        String(row[woIndex]).trim() === String(formData.wo_no).trim()
      );

      const numsSet = new Set<number>();
      const pipeToByBatch = new Map<number, number>();
      clientWoRecords.forEach(row => {
        const raw = row[batchIndex];
        const num = raw != null ? parseInt(String(raw).match(/(\d+)/)?.[1] || '') : NaN;
        if (!isNaN(num) && num > 0) {
          numsSet.add(num);
          if (pipeToIndex !== -1) {
            const pt = parseInt(String(row[pipeToIndex]).replace(/[^\d-]/g, ''));
            pipeToByBatch.set(num, isNaN(pt) ? 0 : pt);
          }
        }
      });

      // Next batch = (max existing) + 1, не заполняем пропуски
      const nums = Array.from(numsSet);
      const maxNum = nums.length ? Math.max(...nums) : 0;
      const nextNum = maxNum + 1;
      const lastFromMax = maxNum > 0 ? (pipeToByBatch.get(maxNum) || 0) : 0;

      return { nextBatchLabel: `Batch # ${nextNum}`, nextBatchNum: nextNum, lastPipeTo: lastFromMax };
    } catch (e) {
      console.error('❌ Error recomputing next batch:', e);
      return { nextBatchLabel: 'Batch # 1', nextBatchNum: 1, lastPipeTo: 0 };
    }
  };

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

  // Update form data when batch is calculated (simplified)
  useEffect(() => {
    if (nextBatch && nextBatch !== formData.batch) {
      console.log('📝 Updating form batch to:', nextBatch);
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
    
    console.log('🚀 Starting tubing form submission...');
    console.log('📋 Form data:', formData);
    console.log('👤 User:', user);
    console.log('🔗 SharePoint service:', sharePointService);
    console.log('🌐 Is connected:', isConnected);
    
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

    if (!isConnected) {
      toast({
        title: "Ошибка",
        description: "SharePoint не подключен. Нажмите 'Load Data' в заголовке",
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

    // Lock: recompute expected batch at submit-time to prevent skipping
    const recomputed = recomputeNextBatchInfo();
    const currentBatchLabel = nextBatch || formData.batch;
    if (!currentBatchLabel) {
      toast({
        title: "Ошибка валидации",
        description: "Batch номер не рассчитан. Попробуйте выбрать клиента и Work Order заново",
        variant: "destructive",
      });
      return;
    }
    if (recomputed.nextBatchLabel !== currentBatchLabel) {
      console.warn('🔒 Batch lock: updating to latest expected value', { expected: recomputed.nextBatchLabel, current: currentBatchLabel });
      setNextBatch(recomputed.nextBatchLabel);
      setLastPipeTo(recomputed.lastPipeTo);
      setFormData(prev => ({ ...prev, batch: recomputed.nextBatchLabel }));
      toast({
        title: "Обновление Batch",
        description: `Номер партии изменился на ${recomputed.nextBatchLabel} для сохранения последовательности. Проверьте и отправьте снова.`,
        variant: "destructive",
      });
      return;
    }

    const batchToUse = currentBatchLabel;

    // Validate inspection quantities
    if (!validateInspectionQuantities()) {
      return;
    }

    setIsLoading(true);
    console.log('📤 Submitting tubing record with batch:', batchToUse);
    console.log('📋 Full form data being sent:', {
      client: formData.client,
      wo_no: formData.wo_no,
      batch: batchToUse,
      diameter: formData.diameter,
      qty: formData.qty,
      pipe_from: formData.pipe_from,
      pipe_to: formData.pipe_to
    });
    
    try {
      console.log('🔥 Calling sharePointService.addTubingRecordToExcel...');
      const success = await sharePointService.addTubingRecordToExcel({
        client: formData.client,
        wo_no: formData.wo_no,
        batch: batchToUse,
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
          title: "❌ Не удалось сохранить в Excel",
          description: (
            <div className="space-y-2">
              <p className="font-bold text-white">
                Возможные причины: файл занят или устарела сессия доступа
              </p>
              <p className="text-sm text-white font-medium">
                ⏳ Подождите 10–20 секунд и попробуйте снова. Если повторяется — нажмите <b>Update Data</b> в заголовке, чтобы обновить сессию, и повторите попытку.
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
    setFormData(prev => {
      const next: any = { ...prev, [field]: value };
      if (field === 'client' && value !== prev.client) {
        // Сброс зависимых полей при смене клиента
        next.wo_no = '';
        next.diameter = '';
        next.batch = '';
        next.pipe_from = '';
        next.pipe_to = '';
        setNextBatch('');
        setAvailableDiameters([]);
        setLastPipeTo(0);
      }
      if (field === 'wo_no' && value !== prev.wo_no) {
        // Сброс зависимых полей при смене WO
        next.diameter = '';
        next.batch = '';
        next.pipe_from = '';
        next.pipe_to = '';
        setNextBatch('');
        setLastPipeTo(0);
      }
      if (field === 'qty' && !value) {
        next.pipe_from = '';
        next.pipe_to = '';
      }
      return next;
    });
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
                      <SelectValue placeholder="Select work order" />
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
                    className="bg-gray-50 border-2"
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
                  <Input
                    id="start_date"
                    type="date"
                    value={formData.start_date}
                    onChange={(e) => handleInputChange("start_date", e.target.value)}
                  />
                </div>

                <div>
                  <Label htmlFor="end_date">End Date</Label>
                  <Input
                    id="end_date"
                    type="date"
                    value={formData.end_date}
                    onChange={(e) => handleInputChange("end_date", e.target.value)}
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
                  <Input
                    id="act_date"
                    type="date"
                    value={formData.act_date}
                    onChange={(e) => handleInputChange("act_date", e.target.value)}
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
