import { useState, useEffect, useMemo } from "react";
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
import { safeLocalStorage } from "@/lib/safe-storage";
import { ConfirmDialog } from "@/components/ui/confirm-dialog";

const stagePriceFields = [
  { key: "rattling_price", label: "Item: 1.7 Rattling_Price" },
  { key: "external_price", label: "Item: 1.1 External_Price" },
  { key: "hydro_price", label: "Item: 1.2 Hydro_Price" },
  { key: "mpi_price", label: "Item: 1.5 MPI_Price" },
  { key: "drift_price", label: "Item: 1.3 Drift_Price" },
  { key: "emi_price", label: "Item: 1.4 EMI_Price" },
  { key: "marking_price", label: "Item: 1.6 Marking_Price" },
] as const;

type StagePriceKey = (typeof stagePriceFields)[number]["key"];
type StagePrices = Record<StagePriceKey, string>;

const createEmptyStagePrices = (): StagePrices => ({
  rattling_price: "",
  external_price: "",
  hydro_price: "",
  mpi_price: "",
  drift_price: "",
  emi_price: "",
  marking_price: "",
});

const sanitizeDecimalInput = (input: string): string => {
  if (!input) return "";
  const normalized = input.replace(/,/g, ".");
  // remove all except digits and dots
  const filtered = normalized.replace(/[^0-9.]/g, "");

  // handle inputs that start with a dot like ".5" -> "0.5"
  if (filtered.startsWith(".") && filtered !== ".") {
    const after = filtered.slice(1).replace(/\./g, "");
    return after ? `0.${after}` : "0.";
  }

  // split by dots to collapse multiples into a single decimal point
  const parts = filtered.split(".");
  if (parts.length === 1) {
    return parts[0];
  }

  const first = parts.shift() ?? "";
  const rest = parts.join("");

  // preserve a single trailing dot while user is typing, e.g. "245." should remain as-is
  const hasTrailingDotOnly = filtered.endsWith(".") && rest.length === 0;
  if (hasTrailingDotOnly) {
    return first + ".";
  }

  return first + (rest ? "." + rest : "");
};

const sanitizeIntegerInput = (input: string): string => input.replace(/[^0-9]/g, "");

export default function WOForm() {
  const navigate = useNavigate();
  const { user } = useAuth();
  const { sharePointService, isConnected, refreshDataInBackground } = useSharePoint();
  const { toast } = useToast();
  const { clients, clientRecords } = useSharePointInstantData();

  const availableClients = useMemo(() => {
    const names = clientRecords.length > 0
      ? clientRecords.map(record => record.name)
      : clients;
    return Array.from(new Set(names.filter(name => name && name.trim()))).sort((a, b) =>
      a.localeCompare(b)
    );
  }, [clientRecords, clients]);

  const [existingWorkOrders, setExistingWorkOrders] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [isConfirmOpen, setIsConfirmOpen] = useState(false);
  const [confirmLines, setConfirmLines] = useState<string[]>([]);
  const [pending, setPending] = useState<{ payload: any; trimmedWo: string } | null>(null);
  const [formData, setFormData] = useState({
    wo_no: "",
    client: "",
    wo_date: "",
    wo_type: "",
    pipe_type: "",
    diameter: "",
    planned_qty: "",
    price_type: "",
    price_per_pipe: "",
    stage_prices: createEmptyStagePrices(),
    transport: "",
    transport_cost: "",
    replacement_price: "",
  });

  useEffect(() => {
    if (!formData.client || !sharePointService) {
      setExistingWorkOrders([]);
      return;
    }

    const loadWorkOrders = async () => {
      try {
        const workOrders = await sharePointService.getWorkOrdersByClient(formData.client);
        setExistingWorkOrders(workOrders);
      } catch (error) {
        console.error("❌ Error loading work orders:", error);
        setExistingWorkOrders([]);
      }
    };

    loadWorkOrders();
  }, [formData.client, sharePointService]);

  // Removed Payer auto-fill logic per requirement: Payer is managed only in Add Clients / Client Edits

  // When Sucker Rod is selected, Diameter is not applicable -> clear it
  useEffect(() => {
    if (formData.pipe_type === "Sucker Rod" && formData.diameter) {
      setFormData(prev => ({ ...prev, diameter: "" }));
    }
  }, [formData.pipe_type, formData.diameter]);

  const handleFieldChange = (field: keyof typeof formData, value: string) => {
    setFormData(prev => ({ ...prev, [field]: value }));
  };

  const handleStagePriceChange = (key: StagePriceKey, value: string) => {
    const sanitized = sanitizeDecimalInput(value);
    setFormData(prev => ({
      ...prev,
      stage_prices: {
        ...prev.stage_prices,
        [key]: sanitized,
      },
    }));
  };

  const resetStagePrices = () => {
    setFormData(prev => ({
      ...prev,
      stage_prices: createEmptyStagePrices(),
    }));
  };

  const validateStagePrices = () => {
    return stagePriceFields.every(field => formData.stage_prices[field.key]);
  };

  const isCouplingReplace = formData.wo_type === "Coupling Replace";
  const isOctgInspection = formData.wo_type === "OCTG Inspection";
  const isStageBasedPricing = formData.price_type === "Stage Based";

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

    if (!formData.client || !formData.wo_no) {
      toast({
        title: "Ошибка валидации",
        description: "Заполните номер Work Order и выберите клиента",
        variant: "destructive",
      });
      return;
    }

    if (!formData.wo_type) {
      toast({
        title: "Ошибка валидации",
        description: "Выберите Type of WO",
        variant: "destructive",
      });
      return;
    }

    if (!formData.wo_date) {
      toast({
        title: "Ошибка валидации",
        description: "Выберите дату Work Order",
        variant: "destructive",
      });
      return;
    }

    if (isOctgInspection) {
      if (!formData.pipe_type) {
        toast({
          title: "Ошибка валидации",
          description: "Выберите Type Of Pipe",
          variant: "destructive",
        });
        return;
      }

      if (formData.pipe_type !== "Sucker Rod" && !formData.diameter) {
        toast({
          title: "Ошибка валидации",
          description: "Выберите Diameter",
          variant: "destructive",
        });
        return;
      }

      if (!formData.planned_qty) {
        toast({
          title: "Ошибка валидации",
          description: "Укажите Planned Qty",
          variant: "destructive",
        });
        return;
      }

      if (!formData.price_type) {
        toast({
          title: "Ошибка валидации",
          description: "Выберите Price Type",
          variant: "destructive",
        });
        return;
      }

      if (formData.price_type === "Fixed" && !formData.price_per_pipe) {
        toast({
          title: "Ошибка валидации",
          description: "Укажите Price for each pipe",
          variant: "destructive",
        });
        return;
      }

      if (formData.price_type === "Stage Based" && !validateStagePrices()) {
        toast({
          title: "Ошибка валидации",
          description: "Заполните цены для всех этапов Stage Based",
          variant: "destructive",
        });
        return;
      }

      if (!formData.transport) {
        toast({
          title: "Ошибка валидации",
          description: "Выберите Transport",
          variant: "destructive",
        });
        return;
      }

      if (formData.transport === "TCC" && !formData.transport_cost) {
        toast({
          title: "Ошибка валидации",
          description: "Укажите Transportation Cost",
          variant: "destructive",
        });
        return;
      }
    }

    if (isCouplingReplace && !formData.replacement_price) {
      toast({
        title: "Ошибка валидации",
        description: "Укажите Price for replacement",
        variant: "destructive",
      });
      return;
    }

    const trimmedWo = formData.wo_no.toString().trim();
    const isDuplicate = existingWorkOrders.some(wo => wo.toString().trim() === trimmedWo);

    if (isDuplicate) {
      toast({
        title: "🚫 Work Order уже существует",
        description: (
          <div className="space-y-2">
            <p className="font-bold text-white">
              Work Order <span className="bg-white text-red-600 px-2 py-1 rounded font-mono font-bold">{trimmedWo}</span> уже существует для клиента <span className="bg-white text-blue-600 px-2 py-1 rounded font-bold">{formData.client}</span>
            </p>
            <p className="text-sm text-white font-medium">💡 Пожалуйста, выберите другой номер Work Order</p>
          </div>
        ),
        variant: "destructive",
        duration: 8000,
      });
      return;
    }

    const payload = {
      wo_no: trimmedWo,
      client: formData.client,
      wo_date: formData.wo_date,
      wo_type: formData.wo_type,
      pipe_type: isCouplingReplace ? "Tubing" : formData.pipe_type,
      type: isCouplingReplace ? "Tubing" : formData.pipe_type,
      diameter: isCouplingReplace ? "" : formData.diameter,
      planned_qty: isOctgInspection ? formData.planned_qty : "",
      price_type: isOctgInspection ? formData.price_type : "",
      price_per_pipe: isOctgInspection && formData.price_type === "Fixed" ? formData.price_per_pipe : "",
      stage_prices: isOctgInspection && isStageBasedPricing ? formData.stage_prices : createEmptyStagePrices(),
      transport: isOctgInspection ? formData.transport : "",
      transport_cost: isOctgInspection && formData.transport === "TCC" ? formData.transport_cost : "",
      replacement_price: isCouplingReplace ? formData.replacement_price : "",
      coupling_replace: isCouplingReplace ? "Yes" : "No",
    };
    // Open in-page confirm dialog
    setConfirmLines([
      `Client: ${formData.client}`,
      `WO No: ${trimmedWo}`,
      `Type of WO: ${formData.wo_type || (isCouplingReplace ? 'Coupling Replace' : 'OCTG Inspection')}`,
    ]);
    setPending({ payload, trimmedWo });
    setIsConfirmOpen(true);
  };

  const doSave = async () => {
    if (!sharePointService || !pending) {
      setIsConfirmOpen(false);
      return;
    }
    const { payload, trimmedWo } = pending;

    setIsLoading(true);
    try {
      const success = await sharePointService.createWorkOrder(payload);

      if (success) {
        toast({
          title: "✅ Work Order создан успешно!",
          description: (
            <div className="space-y-2">
              <p className="font-bold text-white">
                Work Order <span className="bg-white text-green-600 px-2 py-1 rounded font-mono font-bold">{trimmedWo}</span> для клиента <span className="bg-white text-blue-600 px-2 py-1 rounded font-bold">{formData.client}</span>
              </p>
              <p className="text-sm text-white font-medium">🎉 Данные успешно сохранены в SharePoint Excel</p>
            </div>
          ),
          duration: 6000,
        });

        const preservedClient = formData.client;
        const preservedWoType = formData.wo_type;
        const preservedPipeType = preservedWoType === "Coupling Replace" ? "Tubing" : formData.pipe_type;
        const preservedDiameter = preservedWoType === "OCTG Inspection" ? formData.diameter : "";
        const preservedPriceType = preservedWoType === "OCTG Inspection" ? formData.price_type : "";
        // Payer not part of Add Work Order form anymore

        setFormData({
          wo_no: "",
          client: preservedClient,
          wo_date: "",
          wo_type: preservedWoType,
          pipe_type: preservedPipeType,
          diameter: preservedDiameter,
          planned_qty: "",
          price_type: preservedPriceType,
          price_per_pipe: "",
          stage_prices: createEmptyStagePrices(),
          transport: "",
          transport_cost: "",
          replacement_price: "",
        });
        resetStagePrices();

        if (preservedClient) {
          const updatedWorkOrders = await sharePointService.getWorkOrdersByClient(preservedClient);
          setExistingWorkOrders(updatedWorkOrders);
        }

        try {
          if (sharePointService && refreshDataInBackground) {
            safeLocalStorage.removeItem("sharepoint_last_refresh");
            await refreshDataInBackground(sharePointService);
          }
        } catch (e) {
          console.warn("Auto Update Data encountered an error:", e);
        }
      } else {
        toast({
          title: "🔒 SharePoint файл заблокирован",
          description: (
            <div className="space-y-2">
              <p className="font-bold text-white">Кто-то редактирует Excel файл в данный момент</p>
              <p className="text-sm text-white font-medium">⏳ Попробуйте через несколько минут, когда файл освободится</p>
            </div>
          ),
          variant: "destructive",
          duration: 8000,
        });
      }
    } catch (error) {
      console.error("Error creating work order:", error);
      toast({
        title: "Error",
        description: "Failed to create work order. Please try again.",
        variant: "destructive",
      });
    } finally {
      setIsLoading(false);
      setIsConfirmOpen(false);
      setPending(null);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50">
      <Header />
      <div className="container mx-auto px-6 py-8">
        <div className="mb-6">
          <Button
            variant="ghost"
            onClick={() => navigate("/")}
            className="flex items-center gap-2 text-slate-600"
          >
            <ArrowLeft className="w-4 h-4" />
            <span>Back to Dashboard</span>
          </Button>
        </div>

        <Card className="max-w-5xl mx-auto border-2 border-blue-200 rounded-xl shadow-md">
          <CardHeader className="bg-blue-50 border-b">
            <CardTitle className="text-2xl font-bold text-blue-800">Add Work Order</CardTitle>
          </CardHeader>
          <CardContent className="p-6">
            <ConfirmDialog
              open={isConfirmOpen}
              title="Save Work Order?"
              description="Please confirm saving this Work Order to SharePoint"
              lines={confirmLines}
              confirmText="Save"
              cancelText="Cancel"
              onConfirm={doSave}
              onCancel={() => setIsConfirmOpen(false)}
              loading={isLoading}
            />
            <form onSubmit={handleSubmit} className="space-y-8">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-2">
                  <Label htmlFor="client" className="text-sm font-semibold text-gray-700">Client</Label>
                  <Select value={formData.client} onValueChange={(value) => handleFieldChange("client", value)}>
                    <SelectTrigger className="border-2 focus:border-blue-500 h-11">
                      <SelectValue placeholder="Select client" />
                    </SelectTrigger>
                    <SelectContent>
                      {availableClients.map(client => (
                        <SelectItem key={client} value={client}>{client}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="wo_no" className="text-sm font-semibold text-gray-700">Work Order Number</Label>
                  <Input
                    id="wo_no"
                    value={formData.wo_no}
                    onChange={(e) => handleFieldChange("wo_no", sanitizeIntegerInput(e.target.value))}
                    placeholder="Enter WO number"
                    className="border-2 focus:border-blue-500 h-11"
                    inputMode="numeric"
                    required
                  />
                </div>

                <div className="space-y-2">
                  <Label htmlFor="wo_date" className="text-sm font-semibold text-gray-700">Work Order Date</Label>
                  <DateInputField
                    id="wo_date"
                    value={formData.wo_date}
                    onChange={(value) => handleFieldChange("wo_date", value)}
                    className="border-2 focus:border-blue-500 h-11"
                    placeholder="dd/mm/yyyy"
                  />
                </div>

                <div className="space-y-2">
                  <Label htmlFor="wo_type" className="text-sm font-semibold text-gray-700">Type of WO</Label>
                  <Select
                    value={formData.wo_type}
                    onValueChange={(value) => {
                      handleFieldChange("wo_type", value);
                      if (value === "Coupling Replace") {
                        handleFieldChange("pipe_type", "Tubing");
                        handleFieldChange("price_type", "");
                        handleFieldChange("transport", "");
                        handleFieldChange("transport_cost", "");
                        resetStagePrices();
                      } else if (value === "OCTG Inspection") {
                        handleFieldChange("pipe_type", "");
                        handleFieldChange("replacement_price", "");
                      }
                    }}
                  >
                    <SelectTrigger className="border-2 focus:border-blue-500 h-11">
                      <SelectValue placeholder="Select work order type" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="OCTG Inspection">OCTG Inspection</SelectItem>
                      <SelectItem value="Coupling Replace">Coupling Replace</SelectItem>
                    </SelectContent>
                  </Select>
              </div>
            </div>

              {isOctgInspection && (
                <div className="space-y-8">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div className="space-y-2">
                      <Label htmlFor="pipe_type" className="text-sm font-semibold text-gray-700">Type Of Pipe</Label>
                      <Select
                        value={formData.pipe_type}
                        onValueChange={(value) => {
                          handleFieldChange("pipe_type", value);
                          if (value === "Sucker Rod") {
                            handleFieldChange("diameter", "");
                          }
                        }}
                      >
                        <SelectTrigger className="border-2 focus:border-blue-500 h-11">
                          <SelectValue placeholder="Select type of pipe" />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="Tubing">Tubing</SelectItem>
                          <SelectItem value="Sucker Rod">Sucker Rod</SelectItem>
                        </SelectContent>
                      </Select>
                    </div>
                    <div className="space-y-2">
                      <Label htmlFor="diameter" className="text-sm font-semibold text-gray-700">Diameter</Label>
                      <Select
                        value={formData.diameter}
                        onValueChange={(value) => handleFieldChange("diameter", value)}
                        disabled={formData.pipe_type === "Sucker Rod"}
                      >
                        <SelectTrigger className="border-2 focus:border-blue-500 h-11">
                          <SelectValue placeholder={formData.pipe_type === "Sucker Rod" ? "N/A for Sucker Rod" : "Select diameter"} />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value='3 1/2"'>3 1/2"</SelectItem>
                          <SelectItem value='2 7/8"'>2 7/8"</SelectItem>
                        </SelectContent>
                      </Select>
                    </div>
                    <div className="space-y-2">
                      <Label htmlFor="planned_qty" className="text-sm font-semibold text-gray-700">Planned Qty</Label>
                      <Input
                        id="planned_qty"
                        value={formData.planned_qty}
                        onChange={(e) => handleFieldChange("planned_qty", sanitizeIntegerInput(e.target.value))}
                        placeholder="Enter planned quantity"
                        className="border-2 focus:border-blue-500 h-11"
                        inputMode="numeric"
                      />
                    </div>
                    <div className="space-y-2">
                      <Label htmlFor="price_type" className="text-sm font-semibold text-gray-700">Price Type</Label>
                      <Select
                        value={formData.price_type}
                        onValueChange={(value) => {
                          handleFieldChange("price_type", value);
                          handleFieldChange("price_per_pipe", "");
                          resetStagePrices();
                        }}
                      >
                        <SelectTrigger className="border-2 focus:border-blue-500 h-11">
                          <SelectValue placeholder="Select price type" />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="Fixed">Fixed</SelectItem>
                          <SelectItem value="Stage Based">Stage Based</SelectItem>
                        </SelectContent>
                      </Select>
                      <p className="text-xs text-blue-600 font-medium">Используйте точку (.) для десятичных значений</p>
                    </div>
                  </div>

                  {formData.price_type === "Fixed" && (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div className="space-y-2">
                        <Label htmlFor="price_per_pipe" className="text-sm font-semibold text-gray-700">Price for each pipe</Label>
                        <Input
                          id="price_per_pipe"
                          value={formData.price_per_pipe}
                          onChange={(e) => handleFieldChange("price_per_pipe", sanitizeDecimalInput(e.target.value))}
                          placeholder="Enter price per pipe"
                          className="border-2 focus:border-blue-500 h-11"
                          inputMode="decimal"
                        />
                      </div>
                    </div>
                  )}

                  {isStageBasedPricing && (
                    <div className="space-y-4">
                      <div className="border-2 border-dashed border-blue-200 rounded-lg overflow-hidden">
                        {stagePriceFields.map(field => (
                          <div
                            key={field.key}
                            className="grid grid-cols-1 md:grid-cols-2 gap-4 px-4 py-3 border-b border-blue-100 last:border-b-0"
                          >
                            <div className="text-sm font-semibold text-gray-700">{field.label}</div>
                            <Input
                              value={formData.stage_prices[field.key]}
                              onChange={(e) => handleStagePriceChange(field.key, e.target.value)}
                              placeholder="0.00"
                              className="border-2 focus:border-blue-500 h-11"
                              inputMode="decimal"
                            />
                          </div>
                        ))}
                      </div>
                      <p className="text-xs text-blue-600 font-medium">Цены должны содержать только цифры и точку. Запятые автоматически заменяются.</p>
                    </div>
                  )}

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div className="space-y-2">
                      <Label htmlFor="transport" className="text-sm font-semibold text-gray-700">Transport</Label>
                      <Select
                        value={formData.transport}
                        onValueChange={(value) => {
                          handleFieldChange("transport", value);
                          if (value === "Client") {
                            handleFieldChange("transport_cost", "");
                          }
                        }}
                      >
                        <SelectTrigger className="border-2 focus:border-blue-500 h-11">
                          <SelectValue placeholder="Select transport" />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="Client">Client</SelectItem>
                          <SelectItem value="TCC">TCC</SelectItem>
                        </SelectContent>
                      </Select>
                    </div>
                    {formData.transport === "TCC" && (
                      <div className="space-y-2">
                        <Label htmlFor="transport_cost" className="text-sm font-semibold text-gray-700">Transportation Cost</Label>
                        <Input
                          id="transport_cost"
                          value={formData.transport_cost}
                          onChange={(e) => handleFieldChange("transport_cost", sanitizeDecimalInput(e.target.value))}
                          placeholder="Enter transport cost"
                          className="border-2 focus:border-blue-500 h-11"
                          inputMode="decimal"
                        />
                      </div>
                    )}
                  </div>
                </div>
              )}

              {isCouplingReplace && (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-2">
                    <Label className="text-sm font-semibold text-gray-700">Type Of Pipe</Label>
                    <Input
                      value="Tubing"
                      readOnly
                      className="h-11 w-full rounded-md border border-gray-300 bg-gray-100 px-3 text-gray-600 shadow-sm"
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="replacement_price" className="text-sm font-semibold text-gray-700">Price for replacement</Label>
                    <Input
                      id="replacement_price"
                      value={formData.replacement_price}
                      onChange={(e) => handleFieldChange("replacement_price", sanitizeDecimalInput(e.target.value))}
                      placeholder="Enter replacement price"
                      className="border-2 focus:border-blue-500 h-11"
                      inputMode="decimal"
                    />
                  </div>
                </div>
              )}

              <div className="flex justify-end space-x-4 pt-6 border-t-2 border-gray-100">
                <Button type="button" variant="destructive" onClick={() => navigate("/")} className="h-12 px-6">
                  Cancel
                </Button>
                <Button type="submit" className="h-12 px-6 font-semibold" disabled={isLoading || !isConnected}>
                  {isLoading ? "Saving..." : "Save Work Order"}
                </Button>
              </div>
            </form>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
