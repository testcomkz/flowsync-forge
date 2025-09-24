import { useEffect, useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import { ArrowLeft } from "lucide-react";

import { Header } from "@/components/layout/Header";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue
} from "@/components/ui/select";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { DateInputField, toDateInputValue } from "@/components/ui/date-input";
import { useToast } from "@/hooks/use-toast";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useSharePointInstantData } from "@/hooks/useInstantData";

import type { SharePointService } from "@/services/sharePointService";


type StageKey = "rattling" | "external" | "hydro" | "mpi" | "drift" | "emi" | "marking";
type ScrapKey = "rattling" | "external" | "jetting" | "mpi" | "drift" | "emi";

interface WorkOrderRecord {
  id: string;
  client: string;
  wo_no: string;
  type: string;
  diameter: string;
  coupling_replace: string;
  wo_date: string;
  transport: string;
  key_col: string;
  payer: string;
  planned_qty: string;
  originalKey: string;
  originalClient: string;
  originalWo: string;
}

interface StageValues {
  quantities: Partial<Record<StageKey, string>>;
  scrap: Partial<Record<ScrapKey, string>>;
}

interface TubingRecord extends StageValues {
  id: string;
  client: string;
  wo_no: string;
  batch: string;
  status: string;
  diameter: string;
  qty: string;
  pipe_from: string;
  pipe_to: string;
  rack: string;
  arrival_date: string;
  class_1: string;
  class_2: string;
  class_3: string;
  repair: string;
  scrapTotal: string;
  start_date: string;
  end_date: string;
  load_out_date: string;
  act_no_oper: string;
  act_date: string;
  originalClient: string;
  originalWo: string;
  originalBatch: string;
}

const stageMeta: {
  key: StageKey;
  label: string;
  scrapKey?: ScrapKey;
  scrapLabel?: string;
}[] = [
  { key: "rattling", label: "Rattling Qty", scrapKey: "rattling", scrapLabel: "Rattling Scrap Qty" },
  { key: "external", label: "External Qty", scrapKey: "external", scrapLabel: "External Scrap Qty" },
  { key: "hydro", label: "Hydro Qty", scrapKey: "jetting", scrapLabel: "Jetting Scrap Qty" },
  { key: "mpi", label: "MPI Qty", scrapKey: "mpi", scrapLabel: "MPI Scrap Qty" },
  { key: "drift", label: "Drift Qty", scrapKey: "drift", scrapLabel: "Drift Scrap Qty" },
  { key: "emi", label: "EMI Qty", scrapKey: "emi", scrapLabel: "EMI Scrap Qty" },
  { key: "marking", label: "Marking Qty" }
];

const normalize = (value: unknown) => (value === null || value === undefined ? "" : String(value).trim());

const normalizeLower = (value: unknown) => normalize(value).toLowerCase();

const sanitizeNumberString = (value: string) => value.replace(/[^0-9-]/g, "");

const uniqueSorted = (values: string[]) => Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b));

const formatKeyCol = (data: { wo_no: string; client: string; type: string; diameter: string }) => {
  if (!data.wo_no || !data.client || !data.type || !data.diameter) {
    return "";
  }
  return `${data.wo_no} - ${data.client} - ${data.type} - ${data.diameter}`;
};

const computePipeTo = (pipeFrom: string, qty: string) => {
  const parsedFrom = Number.parseInt(sanitizeNumberString(pipeFrom), 10);
  const parsedQty = Number.parseInt(sanitizeNumberString(qty), 10);
  if (Number.isNaN(parsedFrom) || Number.isNaN(parsedQty)) {
    return "";
  }
  const pipeTo = parsedFrom + parsedQty - 1;
  return pipeTo.toString();
};

const stageLabel = (key: StageKey) => stageMeta.find(stage => stage.key === key)?.label ?? key;

const scrapLabel = (key: ScrapKey) =>
  stageMeta.find(stage => stage.scrapKey === key)?.scrapLabel ?? `${key} Scrap Qty`;

const parseWorkOrders = (data: any[]): WorkOrderRecord[] => {
  if (!Array.isArray(data) || data.length < 2) {
    return [];
  }

  const headers = data[0] as unknown[];

  const findIndex = (matcher: (normalized: string, canonical: string) => boolean) =>
    headers.findIndex(header => {
      const normalized = normalizeLower(header);
      const canonical = normalized.replace(/[\s-]+/g, "_").replace(/_{2,}/g, "_");
      return matcher(normalized, canonical);
    });

  const clientIndex = findIndex(header => header.includes("client"));
  const woIndex = findIndex(header => header.includes("wo"));
  const typeIndex = findIndex(header => header.includes("type"));
  const diameterIndex = findIndex(header => header.includes("diameter") || header.includes("диаметр"));
  const couplingIndex = findIndex(header => header.includes("coupling"));
  const dateIndex = findIndex(header => header.includes("date"));
  const transportIndex = findIndex(header => header.includes("transport"));
  const keyIndex = findIndex(header => header.includes("key"));
  const payerIndex = findIndex(header => header.includes("payer") || header.includes("branch"));
  const plannedQtyIndex = findIndex(header => header.includes("qty") || header.includes("quantity"));

  return (data.slice(1) as unknown[][])
    .map((row, rowIndex) => {
      const client = normalize(clientIndex >= 0 ? row[clientIndex] : "");
      const wo_no = normalize(woIndex >= 0 ? row[woIndex] : "");
      if (!client || !wo_no) {
        return null;
      }

      const type = normalize(typeIndex >= 0 ? row[typeIndex] : "");
      const diameter = normalize(diameterIndex >= 0 ? row[diameterIndex] : "");
      const coupling_replace = normalize(couplingIndex >= 0 ? row[couplingIndex] : "");
      const wo_date = toDateInputValue(dateIndex >= 0 ? row[dateIndex] : "");
      const transport = normalize(transportIndex >= 0 ? row[transportIndex] : "");
      const key_col = normalize(keyIndex >= 0 ? row[keyIndex] : `${wo_no} - ${client} - ${type} - ${diameter}`);
      const payer = normalize(payerIndex >= 0 ? row[payerIndex] : "");
      const planned_qty = normalize(plannedQtyIndex >= 0 ? row[plannedQtyIndex] : "");

      return {
        id: `${rowIndex}-${client}-${wo_no}`,
        client,
        wo_no,
        type,
        diameter,
        coupling_replace,
        wo_date,
        transport,
        key_col,
        payer,
        planned_qty,
        originalKey: key_col,
        originalClient: client,
        originalWo: wo_no
      } satisfies WorkOrderRecord;
    })
    .filter((value): value is WorkOrderRecord => Boolean(value));
};

const parseTubingRecords = (data: any[]): TubingRecord[] => {
  if (!Array.isArray(data) || data.length < 2) {
    return [];
  }

  const headers = data[0] as unknown[];

  const normalizeHeader = (header: unknown) => normalizeLower(header);
  const canonicalize = (header: string) => header.replace(/[\s-]+/g, "_").replace(/_{2,}/g, "_");

  const findIndex = (matcher: (normalized: string, canonical: string) => boolean) =>
    headers.findIndex(header => {
      const normalized = normalizeHeader(header);
      const canonical = canonicalize(normalized);
      return matcher(normalized, canonical);
    });

  const clientIndex = findIndex(header => header.includes("client"));
  const woIndex = findIndex(header => header.includes("wo"));
  const batchIndex = findIndex(header => header.includes("batch"));
  const statusIndex = findIndex(header => header.includes("status"));
  const diameterIndex = findIndex(header => header.includes("diameter") || header.includes("диаметр"));
  const qtyIndex = findIndex((header, canonical) =>
    canonical === "qty" || canonical === "quantity" || (canonical.includes("qty") && !canonical.includes("scrap"))
  );
  const pipeFromIndex = findIndex((header, canonical) => canonical.includes("pipe_from"));
  const pipeToIndex = findIndex((header, canonical) => canonical.includes("pipe_to"));
  const rackIndex = findIndex((header, canonical) => canonical.includes("rack"));
  const arrivalDateIndex = findIndex((header, canonical) => canonical.includes("arrival_date"));
  const class1Index = findIndex((header, canonical) => canonical.includes("class_1") || header.includes("class 1"));
  const class2Index = findIndex((header, canonical) => canonical.includes("class_2") || header.includes("class 2"));
  const class3Index = findIndex((header, canonical) => canonical.includes("class_3") || header.includes("class 3"));
  const repairIndex = findIndex(header => header.includes("repair"));
  const scrapIndex = findIndex((header, canonical) => canonical === "scrap" || canonical.endsWith("_scrap"));
  const startDateIndex = findIndex((header, canonical) => canonical.includes("start_date"));
  const endDateIndex = findIndex((header, canonical) => canonical.includes("end_date"));
  const loadOutDateIndex = findIndex((header, canonical) => canonical.includes("load_out_date") || canonical.includes("loadoutdate"));
  const actNoOperIndex = findIndex((header, canonical) => canonical.includes("act_no_oper") || canonical.includes("actnooper"));
  const actDateIndex = findIndex((header, canonical) => canonical.includes("act_date") || canonical.includes("actdate"));

  const rattlingQtyIndex = findIndex((header, canonical) =>
    canonical.includes("rattling_qty") && !canonical.includes("scrap")
  );
  const externalQtyIndex = findIndex((header, canonical) => canonical.includes("external_qty") && !canonical.includes("scrap"));
  const hydroQtyIndex = findIndex((header, canonical) =>
    (canonical.includes("hydro_qty") || canonical.includes("jetting_qty")) && !canonical.includes("scrap")
  );
  const mpiQtyIndex = findIndex((header, canonical) => canonical.includes("mpi_qty") && !canonical.includes("scrap"));
  const driftQtyIndex = findIndex((header, canonical) => canonical.includes("drift_qty") && !canonical.includes("scrap"));
  const emiQtyIndex = findIndex((header, canonical) => canonical.includes("emi_qty") && !canonical.includes("scrap"));
  const markingQtyIndex = findIndex((header, canonical) => canonical.includes("marking_qty"));

  const rattlingScrapIndex = findIndex((header, canonical) => canonical.includes("rattling_scrap"));
  const externalScrapIndex = findIndex((header, canonical) => canonical.includes("external_scrap"));
  const jettingScrapIndex = findIndex((header, canonical) => canonical.includes("jetting_scrap"));
  const mpiScrapIndex = findIndex((header, canonical) => canonical.includes("mpi_scrap"));
  const driftScrapIndex = findIndex((header, canonical) => canonical.includes("drift_scrap"));
  const emiScrapIndex = findIndex((header, canonical) => canonical.includes("emi_scrap"));

  return (data.slice(1) as unknown[][])
    .map((row, rowIndex) => {
      const client = normalize(clientIndex >= 0 ? row[clientIndex] : "");
      const wo_no = normalize(woIndex >= 0 ? row[woIndex] : "");
      const batch = normalize(batchIndex >= 0 ? row[batchIndex] : "");
      if (!client || !wo_no || !batch) {
        return null;
      }

      const status = normalize(statusIndex >= 0 ? row[statusIndex] : "");
      const diameter = normalize(diameterIndex >= 0 ? row[diameterIndex] : "");
      const qty = normalize(qtyIndex >= 0 ? row[qtyIndex] : "");
      const pipe_from = normalize(pipeFromIndex >= 0 ? row[pipeFromIndex] : "");
      const pipe_to = normalize(pipeToIndex >= 0 ? row[pipeToIndex] : "");
      const rack = normalize(rackIndex >= 0 ? row[rackIndex] : "");
      const arrival_date = toDateInputValue(arrivalDateIndex >= 0 ? row[arrivalDateIndex] : "");
      const class_1 = normalize(class1Index >= 0 ? row[class1Index] : "");
      const class_2 = normalize(class2Index >= 0 ? row[class2Index] : "");
      const class_3 = normalize(class3Index >= 0 ? row[class3Index] : "");
      const repair = normalize(repairIndex >= 0 ? row[repairIndex] : "");
      const scrapTotal = normalize(scrapIndex >= 0 ? row[scrapIndex] : "");
      const start_date = toDateInputValue(startDateIndex >= 0 ? row[startDateIndex] : "");
      const end_date = toDateInputValue(endDateIndex >= 0 ? row[endDateIndex] : "");
      const load_out_date = toDateInputValue(loadOutDateIndex >= 0 ? row[loadOutDateIndex] : "");
      const act_no_oper = normalize(actNoOperIndex >= 0 ? row[actNoOperIndex] : "");
      const act_date = toDateInputValue(actDateIndex >= 0 ? row[actDateIndex] : "");

      const quantities: Partial<Record<StageKey, string>> = {
        rattling: normalize(rattlingQtyIndex >= 0 ? row[rattlingQtyIndex] : ""),
        external: normalize(externalQtyIndex >= 0 ? row[externalQtyIndex] : ""),
        hydro: normalize(hydroQtyIndex >= 0 ? row[hydroQtyIndex] : ""),
        mpi: normalize(mpiQtyIndex >= 0 ? row[mpiQtyIndex] : ""),
        drift: normalize(driftQtyIndex >= 0 ? row[driftQtyIndex] : ""),
        emi: normalize(emiQtyIndex >= 0 ? row[emiQtyIndex] : ""),
        marking: normalize(markingQtyIndex >= 0 ? row[markingQtyIndex] : "")
      };

      const scrap: Partial<Record<ScrapKey, string>> = {
        rattling: normalize(rattlingScrapIndex >= 0 ? row[rattlingScrapIndex] : ""),
        external: normalize(externalScrapIndex >= 0 ? row[externalScrapIndex] : ""),
        jetting: normalize(jettingScrapIndex >= 0 ? row[jettingScrapIndex] : ""),
        mpi: normalize(mpiScrapIndex >= 0 ? row[mpiScrapIndex] : ""),
        drift: normalize(driftScrapIndex >= 0 ? row[driftScrapIndex] : ""),
        emi: normalize(emiScrapIndex >= 0 ? row[emiScrapIndex] : "")
      };

      return {
        id: `${rowIndex}-${client}-${wo_no}-${batch}`,
        client,
        wo_no,
        batch,
        status,
        diameter,
        qty,
        pipe_from,
        pipe_to,
        rack,
        arrival_date,
        class_1,
        class_2,
        class_3,
        repair,
        scrapTotal,
        start_date,
        end_date,
        load_out_date,
        act_no_oper,
        act_date,
        quantities,
        scrap,
        originalClient: client,
        originalWo: wo_no,
        originalBatch: batch
      } satisfies TubingRecord;
    })
    .filter((value): value is TubingRecord => Boolean(value));
};


type ToastFn = ReturnType<typeof useToast>["toast"];

function WorkOrderEditSection({
  records,
  sharePointService,
  isConnected,
  refreshData,
  toast
}: {
  records: WorkOrderRecord[];
  sharePointService: SharePointService | null;
  isConnected: boolean;
  refreshData: ((service: SharePointService) => Promise<void>) | null;
  toast: ToastFn;
}) {
  const clients = useMemo(() => uniqueSorted(records.map(record => record.client)), [records]);

  const [selectedClient, setSelectedClient] = useState<string>("");
  const [selectedWorkOrderId, setSelectedWorkOrderId] = useState<string>("");
  const [formData, setFormData] = useState({
    client: "",
    wo_no: "",
    type: "",
    diameter: "",
    coupling_replace: "",
    wo_date: "",
    transport: "",
    key_col: "",
    payer: "",
    planned_qty: "",
    originalKey: "",
    originalClient: "",
    originalWo: ""
  });
  const [isSaving, setIsSaving] = useState(false);

  const workOrdersForClient = useMemo(
    () => records.filter(record => record.client === selectedClient),
    [records, selectedClient]
  );

  useEffect(() => {
    setSelectedWorkOrderId("");
    setFormData(prev => ({ ...prev, client: selectedClient || "" }));
  }, [selectedClient]);

  useEffect(() => {
    if (!selectedWorkOrderId) {
      setFormData(prev => ({
        ...prev,
        wo_no: "",
        type: "",
        diameter: "",
        coupling_replace: "",
        wo_date: "",
        transport: "",
        key_col: "",
        payer: "",
        planned_qty: "",
        originalKey: "",
        originalClient: selectedClient,
        originalWo: ""
      }));
      return;
    }

    const record = workOrdersForClient.find(item => item.id === selectedWorkOrderId);
    if (!record) {
      return;
    }

    setFormData({
      client: record.client,
      wo_no: record.wo_no,
      type: record.type,
      diameter: record.diameter,
      coupling_replace: record.coupling_replace,
      wo_date: record.wo_date,
      transport: record.transport,
      key_col: record.key_col,
      payer: record.payer,
      planned_qty: record.planned_qty,
      originalKey: record.originalKey,
      originalClient: record.originalClient,
      originalWo: record.originalWo
    });
  }, [selectedWorkOrderId, workOrdersForClient, selectedClient]);

  const handleInputChange = (field: keyof typeof formData, value: string) => {
    setFormData(prev => {
      const next = { ...prev, [field]: value };
      if (["client", "wo_no", "type", "diameter"].includes(field)) {
        next.key_col = formatKeyCol({
          client: next.client,
          wo_no: next.wo_no,
          type: next.type,
          diameter: next.diameter
        });
      }
      return next;
    });
  };

  const handleUpdate = async () => {
    if (!sharePointService || !isConnected) {
      toast({
        title: "SharePoint not connected",
        description: "Connect to SharePoint before updating records.",
        variant: "destructive"
      });
      return;
    }

    if (!formData.client || !formData.wo_no || !selectedWorkOrderId) {
      toast({
        title: "Validation error",
        description: "Select a Work Order and fill in required fields before updating.",
        variant: "destructive"
      });
      return;
    }

    setIsSaving(true);
    try {
      const success = await sharePointService.updateWorkOrder({
        originalKey: formData.originalKey,
        originalClient: formData.originalClient,
        originalWo: formData.originalWo,
        client: formData.client,
        wo_no: formData.wo_no,
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
          title: "Work Order updated",
          description: `${formData.wo_no} for ${formData.client} saved successfully.`
        });
        if (refreshData && sharePointService) {
          await refreshData(sharePointService);
        }
      } else {
        toast({
          title: "Update failed",
          description: "Unable to update Work Order. Please try again.",
          variant: "destructive"
        });
      }
    } catch (error) {
      console.error("Failed to update work order", error);
      toast({
        title: "Update failed",
        description: "Unexpected error occurred while updating Work Order.",
        variant: "destructive"
      });
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <Card className="border-2 shadow-md">
      <CardHeader className="border-b bg-slate-50">
        <CardTitle className="text-xl font-semibold">Work Order Edit</CardTitle>
      </CardHeader>
      <CardContent className="space-y-6 p-6">
        <div className="grid gap-4 md:grid-cols-2">
          <div className="space-y-2">
            <Label htmlFor="wo_client">Client</Label>
            <Select value={selectedClient} onValueChange={setSelectedClient}>
              <SelectTrigger id="wo_client">
                <SelectValue placeholder="Select client" />
              </SelectTrigger>
              <SelectContent>
                {clients.map(client => (
                  <SelectItem key={client} value={client}>
                    {client}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>

          <div className="space-y-2">
            <Label htmlFor="wo_selector">Work Order</Label>
            <Select
              value={selectedWorkOrderId}
              onValueChange={setSelectedWorkOrderId}
              disabled={!selectedClient || workOrdersForClient.length === 0}
            >
              <SelectTrigger id="wo_selector">
                <SelectValue placeholder="Select Work Order" />
              </SelectTrigger>
              <SelectContent>
                {workOrdersForClient.map(record => (
                  <SelectItem key={record.id} value={record.id}>
                    {record.wo_no} · {record.type || "Type"} · {record.diameter || "Diameter"}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </div>

        {selectedWorkOrderId ? (
          <div className="grid gap-4 md:grid-cols-2">
            <div className="space-y-2">
              <Label htmlFor="wo_no">Work Order Number</Label>
              <Input
                id="wo_no"
                value={formData.wo_no}
                onChange={event => handleInputChange("wo_no", event.target.value)}
                placeholder="Enter Work Order number"
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="wo_type">Type</Label>
              <Input
                id="wo_type"
                value={formData.type}
                onChange={event => handleInputChange("type", event.target.value)}
                placeholder="Enter type"
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="wo_diameter">Diameter</Label>
              <Input
                id="wo_diameter"
                value={formData.diameter}
                onChange={event => handleInputChange("diameter", event.target.value)}
                placeholder="Enter diameter"
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="wo_coupling">Coupling Replace</Label>
              <Input
                id="wo_coupling"
                value={formData.coupling_replace}
                onChange={event => handleInputChange("coupling_replace", event.target.value)}
                placeholder="Enter coupling details"
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="wo_date">Work Order Date</Label>
              <DateInputField
                id="wo_date"
                value={formData.wo_date}
                onChange={value => handleInputChange("wo_date", value)}
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="wo_transport">Transport</Label>
              <Input
                id="wo_transport"
                value={formData.transport}
                onChange={event => handleInputChange("transport", event.target.value)}
                placeholder="Enter transport"
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="wo_payer">Branch / Payer</Label>
              <Input
                id="wo_payer"
                value={formData.payer}
                onChange={event => handleInputChange("payer", event.target.value)}
                placeholder="Enter branch or payer"
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="wo_qty">Planned Quantity</Label>
              <Input
                id="wo_qty"
                value={formData.planned_qty}
                onChange={event => handleInputChange("planned_qty", event.target.value)}
                placeholder="Enter quantity"
              />
            </div>
            <div className="space-y-2 md:col-span-2">
              <Label htmlFor="wo_key">Key Column (auto)</Label>
              <Input id="wo_key" value={formData.key_col} disabled />
            </div>
          </div>
        ) : (
          <div className="rounded-lg border border-dashed border-slate-300 bg-slate-50 p-6 text-center text-sm text-slate-500">
            Select client and Work Order to load existing data.
          </div>
        )}

        <div className="flex justify-end">
          <Button onClick={handleUpdate} disabled={!selectedWorkOrderId || isSaving}>
            {isSaving ? "Updating..." : "Update Work Order"}
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}

function TubingEditSection({
  records,
  sharePointService,
  isConnected,
  refreshData,
  toast
}: {
  records: TubingRecord[];
  sharePointService: SharePointService | null;
  isConnected: boolean;
  refreshData: ((service: SharePointService) => Promise<void>) | null;
  toast: ToastFn;
}) {
  const arrivedRecords = useMemo(
    () => records.filter(record => normalizeLower(record.status) === "arrived"),
    [records]
  );

  const clients = useMemo(
    () => uniqueSorted(arrivedRecords.map(record => record.client)),
    [arrivedRecords]
  );

  const [selectedClient, setSelectedClient] = useState("");
  const [selectedWorkOrderId, setSelectedWorkOrderId] = useState("");
  const [selectedRecordId, setSelectedRecordId] = useState("");
  const [formData, setFormData] = useState({
    client: "",
    wo_no: "",
    batch: "",
    diameter: "",
    qty: "",
    pipe_from: "",
    pipe_to: "",
    rack: "",
    arrival_date: "",
    originalClient: "",
    originalWo: "",
    originalBatch: ""
  });
  const [isSaving, setIsSaving] = useState(false);

  const workOrdersForClient = useMemo(
    () => arrivedRecords.filter(record => record.client === selectedClient),
    [arrivedRecords, selectedClient]
  );

  const batchesForWorkOrder = useMemo(
    () => workOrdersForClient.filter(record => record.wo_no === selectedWorkOrderId),
    [workOrdersForClient, selectedWorkOrderId]
  );

  useEffect(() => {
    setSelectedWorkOrderId("");
    setSelectedRecordId("");
    setFormData(prev => ({ ...prev, client: selectedClient || "", wo_no: "" }));
  }, [selectedClient]);

  useEffect(() => {
    setSelectedRecordId("");
    setFormData(prev => ({ ...prev, wo_no: selectedWorkOrderId || "" }));
  }, [selectedWorkOrderId]);

  useEffect(() => {
    if (!selectedRecordId) {
      return;
    }
    const record = batchesForWorkOrder.find(item => item.id === selectedRecordId);
    if (!record) {
      return;
    }
    setFormData({
      client: record.client,
      wo_no: record.wo_no,
      batch: record.batch,
      diameter: record.diameter,
      qty: record.qty,
      pipe_from: record.pipe_from,
      pipe_to: record.pipe_to,
      rack: record.rack,
      arrival_date: record.arrival_date,
      originalClient: record.originalClient,
      originalWo: record.originalWo,
      originalBatch: record.originalBatch
    });
  }, [selectedRecordId, batchesForWorkOrder]);

  const handleInputChange = (field: keyof typeof formData, value: string) => {
    setFormData(prev => {
      const next = { ...prev, [field]: value };
      if (field === "qty" || field === "pipe_from") {
        const computed = computePipeTo(field === "qty" ? next.pipe_from : value, field === "pipe_from" ? next.qty : value);
        if (computed) {
          next.pipe_to = computed;
        }
      }
      return next;
    });
  };

  const handleUpdate = async () => {
    if (!sharePointService || !isConnected) {
      toast({
        title: "SharePoint not connected",
        description: "Connect to SharePoint before updating records.",
        variant: "destructive"
      });
      return;
    }

    if (!selectedRecordId) {
      toast({
        title: "No batch selected",
        description: "Choose an Arrived batch to update.",
        variant: "destructive"
      });
      return;
    }

    setIsSaving(true);
    try {
      const success = await sharePointService.updateTubingRecord({
        originalClient: formData.originalClient,
        originalWo: formData.originalWo,
        originalBatch: formData.originalBatch,
        client: formData.client,
        wo_no: formData.wo_no,
        batch: formData.batch,
        diameter: formData.diameter,
        qty: formData.qty,
        pipe_from: formData.pipe_from,
        pipe_to: formData.pipe_to,
        rack: formData.rack,
        arrival_date: formData.arrival_date,
        status: "Arrived"
      });

      if (success) {
        toast({
          title: "Tubing batch updated",
          description: `${formData.batch} saved successfully.`
        });
        if (refreshData && sharePointService) {
          await refreshData(sharePointService);
        }
      } else {
        toast({
          title: "Update failed",
          description: "Unable to update tubing record. Please try again.",
          variant: "destructive"
        });
      }
    } catch (error) {
      console.error("Failed to update tubing record", error);
      toast({
        title: "Update failed",
        description: "Unexpected error occurred while updating tubing record.",
        variant: "destructive"
      });
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <Card className="border-2 shadow-md">
      <CardHeader className="border-b bg-emerald-50">
        <CardTitle className="text-xl font-semibold text-emerald-900">Tubing Registry Edit</CardTitle>
      </CardHeader>
      <CardContent className="space-y-6 p-6">
        <div className="grid gap-4 md:grid-cols-3">
          <div className="space-y-2">
            <Label>Client</Label>
            <Select value={selectedClient} onValueChange={setSelectedClient}>
              <SelectTrigger>
                <SelectValue placeholder="Select client" />
              </SelectTrigger>
              <SelectContent>
                {clients.map(client => (
                  <SelectItem key={client} value={client}>
                    {client}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <div className="space-y-2">
            <Label>Work Order</Label>
            <Select
              value={selectedWorkOrderId}
              onValueChange={setSelectedWorkOrderId}
              disabled={!selectedClient}
            >
              <SelectTrigger>
                <SelectValue placeholder="Select Work Order" />
              </SelectTrigger>
              <SelectContent>
                {workOrdersForClient.map(record => (
                  <SelectItem key={`${record.id}-wo`} value={record.wo_no}>
                    {record.wo_no}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <div className="space-y-2">
            <Label>Batch</Label>
            <Select
              value={selectedRecordId}
              onValueChange={setSelectedRecordId}
              disabled={!selectedWorkOrderId}
            >
              <SelectTrigger>
                <SelectValue placeholder="Select batch" />
              </SelectTrigger>
              <SelectContent>
                {batchesForWorkOrder.map(record => (
                  <SelectItem key={record.id} value={record.id}>
                    {record.batch} · Qty {record.qty || ""}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </div>

        {selectedRecordId ? (
          <div className="grid gap-4 md:grid-cols-2">
            <div className="space-y-2">
              <Label htmlFor="tubing_client">Client</Label>
              <Input
                id="tubing_client"
                value={formData.client}
                onChange={event => handleInputChange("client", event.target.value)}
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="tubing_wo">Work Order</Label>
              <Input
                id="tubing_wo"
                value={formData.wo_no}
                onChange={event => handleInputChange("wo_no", event.target.value)}
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="tubing_batch">Batch</Label>
              <Input
                id="tubing_batch"
                value={formData.batch}
                onChange={event => handleInputChange("batch", event.target.value)}
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="tubing_diameter">Diameter</Label>
              <Input
                id="tubing_diameter"
                value={formData.diameter}
                onChange={event => handleInputChange("diameter", event.target.value)}
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="tubing_qty">Quantity</Label>
              <Input
                id="tubing_qty"
                value={formData.qty}
                onChange={event => handleInputChange("qty", event.target.value)}
                inputMode="numeric"
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="tubing_rack">Rack</Label>
              <Input
                id="tubing_rack"
                value={formData.rack}
                onChange={event => handleInputChange("rack", event.target.value)}
              />
            </div>
            <div className="space-y-2">
              <Label htmlFor="tubing_arrival">Arrival Date</Label>
              <DateInputField
                id="tubing_arrival"
                value={formData.arrival_date}
                onChange={value => handleInputChange("arrival_date", value)}
              />
            </div>
            <div className="space-y-2 md:col-span-2 grid grid-cols-2 gap-4">
              <div>
                <Label htmlFor="tubing_from">Pipe From</Label>
                <Input id="tubing_from" value={formData.pipe_from} disabled />
              </div>
              <div>
                <Label htmlFor="tubing_to">Pipe To</Label>
                <Input id="tubing_to" value={formData.pipe_to} disabled />
              </div>
            </div>
          </div>
        ) : (
          <div className="rounded-lg border border-dashed border-emerald-300 bg-emerald-50 p-6 text-center text-sm text-emerald-800">
            Select a batch with status Arrived to edit its data.
          </div>
        )}

        <div className="flex justify-end">
          <Button onClick={handleUpdate} disabled={!selectedRecordId || isSaving}>
            {isSaving ? "Updating..." : "Update Tubing Record"}
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}

function InspectionEditSection({
  records,
  sharePointService,
  isConnected,
  refreshData,
  toast
}: {
  records: TubingRecord[];
  sharePointService: SharePointService | null;
  isConnected: boolean;
  refreshData: ((service: SharePointService) => Promise<void>) | null;
  toast: ToastFn;
}) {
  const eligibleRecords = useMemo(
    () =>
      records.filter(record => {
        const status = normalizeLower(record.status);
        return status === "arrived" || status === "inspection done";
      }),
    [records]
  );

  const clients = useMemo(
    () => uniqueSorted(eligibleRecords.map(record => record.client)),
    [eligibleRecords]
  );

  const [selectedClient, setSelectedClient] = useState("");
  const [selectedWorkOrderId, setSelectedWorkOrderId] = useState("");
  const [selectedRecordId, setSelectedRecordId] = useState("");
  const [formData, setFormData] = useState({
    client: "",
    wo_no: "",
    batch: "",
    status: "",
    class_1: "",
    class_2: "",
    class_3: "",
    repair: "",
    scrapTotal: "",
    start_date: "",
    end_date: "",
    quantities: {} as Partial<Record<StageKey, string>>,
    scrap: {} as Partial<Record<ScrapKey, string>>,
    originalClient: "",
    originalWo: "",
    originalBatch: ""
  });
  const [isSaving, setIsSaving] = useState(false);

  const workOrdersForClient = useMemo(
    () => eligibleRecords.filter(record => record.client === selectedClient),
    [eligibleRecords, selectedClient]
  );

  const batchesForWorkOrder = useMemo(
    () => workOrdersForClient.filter(record => record.wo_no === selectedWorkOrderId),
    [workOrdersForClient, selectedWorkOrderId]
  );

  useEffect(() => {
    setSelectedWorkOrderId("");
    setSelectedRecordId("");
  }, [selectedClient]);

  useEffect(() => {
    setSelectedRecordId("");
  }, [selectedWorkOrderId]);

  useEffect(() => {
    if (!selectedRecordId) {
      return;
    }
    const record = batchesForWorkOrder.find(item => item.id === selectedRecordId);
    if (!record) {
      return;
    }
    setFormData({
      client: record.client,
      wo_no: record.wo_no,
      batch: record.batch,
      status: record.status || "Inspection Done",
      class_1: record.class_1,
      class_2: record.class_2,
      class_3: record.class_3,
      repair: record.repair,
      scrapTotal: record.scrapTotal,
      start_date: record.start_date,
      end_date: record.end_date,
      quantities: { ...record.quantities },
      scrap: { ...record.scrap },
      originalClient: record.originalClient,
      originalWo: record.originalWo,
      originalBatch: record.originalBatch
    });
  }, [selectedRecordId, batchesForWorkOrder]);

  const handleQuantityChange = (key: StageKey, value: string) => {
    setFormData(prev => ({
      ...prev,
      quantities: { ...prev.quantities, [key]: sanitizeNumberString(value) }
    }));
  };

  const handleScrapChange = (key: ScrapKey, value: string) => {
    setFormData(prev => ({
      ...prev,
      scrap: { ...prev.scrap, [key]: sanitizeNumberString(value) }
    }));
  };

  const handleInputChange = (field: keyof typeof formData, value: string) => {
    setFormData(prev => ({ ...prev, [field]: value }));
  };

  const handleUpdate = async () => {
    if (!sharePointService || !isConnected) {
      toast({
        title: "SharePoint not connected",
        description: "Connect to SharePoint before updating records.",
        variant: "destructive"
      });
      return;
    }

    if (!selectedRecordId) {
      toast({
        title: "No batch selected",
        description: "Choose a batch to edit inspection data.",
        variant: "destructive"
      });
      return;
    }

    setIsSaving(true);
    try {
      const success = await sharePointService.updateTubingInspectionData({
        client: formData.client,
        wo_no: formData.wo_no,
        batch: formData.batch,
        class_1: formData.class_1,
        class_2: formData.class_2,
        class_3: formData.class_3,
        repair: formData.repair,
        scrap: formData.scrapTotal,
        start_date: formData.start_date,
        end_date: formData.end_date,
        rattling_qty: Number.parseInt(formData.quantities.rattling || "0", 10) || 0,
        external_qty: Number.parseInt(formData.quantities.external || "0", 10) || 0,
        hydro_qty: Number.parseInt(formData.quantities.hydro || "0", 10) || 0,
        mpi_qty: Number.parseInt(formData.quantities.mpi || "0", 10) || 0,
        drift_qty: Number.parseInt(formData.quantities.drift || "0", 10) || 0,
        emi_qty: Number.parseInt(formData.quantities.emi || "0", 10) || 0,
        marking_qty: Number.parseInt(formData.quantities.marking || "0", 10) || 0,
        status: formData.status || "Inspection Done",
        rattling_scrap_qty: Number.parseInt(formData.scrap.rattling || "0", 10) || 0,
        external_scrap_qty: Number.parseInt(formData.scrap.external || "0", 10) || 0,
        jetting_scrap_qty: Number.parseInt(formData.scrap.jetting || "0", 10) || 0,
        mpi_scrap_qty: Number.parseInt(formData.scrap.mpi || "0", 10) || 0,
        drift_scrap_qty: Number.parseInt(formData.scrap.drift || "0", 10) || 0,
        emi_scrap_qty: Number.parseInt(formData.scrap.emi || "0", 10) || 0,
        originalClient: formData.originalClient,
        originalWo: formData.originalWo,
        originalBatch: formData.originalBatch
      });

      if (success) {
        toast({
          title: "Inspection data updated",
          description: `${formData.batch} saved successfully.`
        });
        if (refreshData && sharePointService) {
          await refreshData(sharePointService);
        }
      } else {
        toast({
          title: "Update failed",
          description: "Unable to update inspection data. Please try again.",
          variant: "destructive"
        });
      }
    } catch (error) {
      console.error("Failed to update inspection data", error);
      toast({
        title: "Update failed",
        description: "Unexpected error occurred while updating inspection data.",
        variant: "destructive"
      });
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <Card className="border-2 shadow-md">
      <CardHeader className="border-b bg-blue-50">
        <CardTitle className="text-xl font-semibold text-blue-900">Inspection Edit</CardTitle>
      </CardHeader>
      <CardContent className="space-y-6 p-6">
        <div className="grid gap-4 md:grid-cols-3">
          <div className="space-y-2">
            <Label>Client</Label>
            <Select value={selectedClient} onValueChange={setSelectedClient}>
              <SelectTrigger>
                <SelectValue placeholder="Select client" />
              </SelectTrigger>
              <SelectContent>
                {clients.map(client => (
                  <SelectItem key={client} value={client}>
                    {client}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <div className="space-y-2">
            <Label>Work Order</Label>
            <Select
              value={selectedWorkOrderId}
              onValueChange={setSelectedWorkOrderId}
              disabled={!selectedClient}
            >
              <SelectTrigger>
                <SelectValue placeholder="Select Work Order" />
              </SelectTrigger>
              <SelectContent>
                {workOrdersForClient.map(record => (
                  <SelectItem key={`${record.id}-inspection`} value={record.wo_no}>
                    {record.wo_no}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <div className="space-y-2">
            <Label>Batch</Label>
            <Select
              value={selectedRecordId}
              onValueChange={setSelectedRecordId}
              disabled={!selectedWorkOrderId}
            >
              <SelectTrigger>
                <SelectValue placeholder="Select batch" />
              </SelectTrigger>
              <SelectContent>
                {batchesForWorkOrder.map(record => (
                  <SelectItem key={record.id} value={record.id}>
                    {record.batch} · {record.status || "Arrived"}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </div>

        {selectedRecordId ? (
          <div className="space-y-6">
            <div className="grid gap-4 md:grid-cols-2">
              <div className="space-y-2">
                <Label>Client</Label>
                <Input
                  value={formData.client}
                  onChange={event => handleInputChange("client", event.target.value)}
                />
              </div>
              <div className="space-y-2">
                <Label>Work Order</Label>
                <Input
                  value={formData.wo_no}
                  onChange={event => handleInputChange("wo_no", event.target.value)}
                />
              </div>
              <div className="space-y-2">
                <Label>Batch</Label>
                <Input
                  value={formData.batch}
                  onChange={event => handleInputChange("batch", event.target.value)}
                />
              </div>
              <div className="space-y-2">
                <Label>Status</Label>
                <Select
                  value={formData.status || "Inspection Done"}
                  onValueChange={value => handleInputChange("status", value)}
                >
                  <SelectTrigger>
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="Inspection Done">Inspection Done</SelectItem>
                    <SelectItem value="Arrived">Arrived</SelectItem>
                  </SelectContent>
                </Select>
              </div>
              <div className="space-y-2">
                <Label>Start Date</Label>
                <DateInputField
                  value={formData.start_date}
                  onChange={value => handleInputChange("start_date", value)}
                />
              </div>
              <div className="space-y-2">
                <Label>End Date</Label>
                <DateInputField
                  value={formData.end_date}
                  onChange={value => handleInputChange("end_date", value)}
                />
              </div>
              <div className="space-y-2">
                <Label>Class 1</Label>
                <Input
                  value={formData.class_1}
                  onChange={event => handleInputChange("class_1", event.target.value)}
                />
              </div>
              <div className="space-y-2">
                <Label>Class 2</Label>
                <Input
                  value={formData.class_2}
                  onChange={event => handleInputChange("class_2", event.target.value)}
                />
              </div>
              <div className="space-y-2">
                <Label>Class 3</Label>
                <Input
                  value={formData.class_3}
                  onChange={event => handleInputChange("class_3", event.target.value)}
                />
              </div>
              <div className="space-y-2">
                <Label>Repair</Label>
                <Input
                  value={formData.repair}
                  onChange={event => handleInputChange("repair", sanitizeNumberString(event.target.value))}
                  inputMode="numeric"
                  placeholder="0"
                />
              </div>
              <div className="space-y-2">
                <Label>Scrap</Label>
                <Input
                  value={formData.scrapTotal}
                  onChange={event => handleInputChange("scrapTotal", sanitizeNumberString(event.target.value))}
                  inputMode="numeric"
                  placeholder="0"
                />
              </div>
            </div>

            <div className="rounded-lg border border-blue-200">
              <div className="border-b bg-blue-50 px-4 py-2 font-semibold text-blue-900">
                Inspection Stages
              </div>
              <div className="grid gap-4 p-4 md:grid-cols-2">
                {stageMeta.map(stage => (
                  <div key={stage.key} className="space-y-2">
                    <div className="flex gap-2">
                      <div className="flex-1">
                        <Label>{stageLabel(stage.key)}</Label>
                        <Input
                          value={formData.quantities[stage.key] ?? ""}
                          onChange={event => handleQuantityChange(stage.key, event.target.value)}
                          inputMode="numeric"
                          placeholder="0"
                        />
                      </div>
                      {stage.scrapKey ? (
                        <div className="flex-1">
                          <Label>{scrapLabel(stage.scrapKey)}</Label>
                          <Input
                            value={formData.scrap[stage.scrapKey] ?? ""}
                            onChange={event => handleScrapChange(stage.scrapKey!, event.target.value)}
                            inputMode="numeric"
                            placeholder="0"
                          />
                        </div>
                      ) : null}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        ) : (
          <div className="rounded-lg border border-dashed border-blue-300 bg-blue-50 p-6 text-center text-sm text-blue-900">
            Select a batch with status Arrived or Inspection Done to view inspection details.
          </div>
        )}

        <div className="flex justify-end">
          <Button onClick={handleUpdate} disabled={!selectedRecordId || isSaving}>
            {isSaving ? "Updating..." : "Update Inspection"}
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}

function LoadOutEditSection({
  records,
  sharePointService,
  isConnected,
  refreshData,
  toast
}: {
  records: TubingRecord[];
  sharePointService: SharePointService | null;
  isConnected: boolean;
  refreshData: ((service: SharePointService) => Promise<void>) | null;
  toast: ToastFn;
}) {
  const eligibleRecords = useMemo(
    () =>
      records.filter(record => {
        const status = normalizeLower(record.status);
        return status === "completed" || status === "inspection done";
      }),
    [records]
  );

  const clients = useMemo(
    () => uniqueSorted(eligibleRecords.map(record => record.client)),
    [eligibleRecords]
  );

  const [selectedClient, setSelectedClient] = useState("");
  const [selectedWorkOrderId, setSelectedWorkOrderId] = useState("");
  const [selectedRecordId, setSelectedRecordId] = useState("");
  const [formData, setFormData] = useState({
    client: "",
    wo_no: "",
    batch: "",
    status: "",
    load_out_date: "",
    act_no_oper: "",
    act_date: "",
    originalClient: "",
    originalWo: "",
    originalBatch: ""
  });
  const [isSaving, setIsSaving] = useState(false);

  const workOrdersForClient = useMemo(
    () => eligibleRecords.filter(record => record.client === selectedClient),
    [eligibleRecords, selectedClient]
  );

  const batchesForWorkOrder = useMemo(
    () => workOrdersForClient.filter(record => record.wo_no === selectedWorkOrderId),
    [workOrdersForClient, selectedWorkOrderId]
  );

  useEffect(() => {
    setSelectedWorkOrderId("");
    setSelectedRecordId("");
  }, [selectedClient]);

  useEffect(() => {
    setSelectedRecordId("");
  }, [selectedWorkOrderId]);

  useEffect(() => {
    if (!selectedRecordId) {
      return;
    }
    const record = batchesForWorkOrder.find(item => item.id === selectedRecordId);
    if (!record) {
      return;
    }
    setFormData({
      client: record.client,
      wo_no: record.wo_no,
      batch: record.batch,
      status: record.status || "Completed",
      load_out_date: record.load_out_date,
      act_no_oper: record.act_no_oper,
      act_date: record.act_date,
      originalClient: record.originalClient,
      originalWo: record.originalWo,
      originalBatch: record.originalBatch
    });
  }, [selectedRecordId, batchesForWorkOrder]);

  const handleInputChange = (field: keyof typeof formData, value: string) => {
    setFormData(prev => ({ ...prev, [field]: value }));
  };

  const handleUpdate = async () => {
    if (!sharePointService || !isConnected) {
      toast({
        title: "SharePoint not connected",
        description: "Connect to SharePoint before updating records.",
        variant: "destructive"
      });
      return;
    }

    if (!selectedRecordId) {
      toast({
        title: "No batch selected",
        description: "Choose a batch to edit load out data.",
        variant: "destructive"
      });
      return;
    }

    setIsSaving(true);
    try {
      const success = await sharePointService.updateLoadOutData({
        client: formData.client,
        wo_no: formData.wo_no,
        batch: formData.batch,
        status: formData.status,
        load_out_date: formData.load_out_date,
        act_no_oper: formData.act_no_oper,
        act_date: formData.act_date,
        originalClient: formData.originalClient,
        originalWo: formData.originalWo,
        originalBatch: formData.originalBatch
      });

      if (success) {
        toast({
          title: "Load Out updated",
          description: `${formData.batch} saved successfully.`
        });
        if (refreshData && sharePointService) {
          await refreshData(sharePointService);
        }
      } else {
        toast({
          title: "Update failed",
          description: "Unable to update load out data. Please try again.",
          variant: "destructive"
        });
      }
    } catch (error) {
      console.error("Failed to update load out", error);
      toast({
        title: "Update failed",
        description: "Unexpected error occurred while updating load out data.",
        variant: "destructive"
      });
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <Card className="border-2 shadow-md">
      <CardHeader className="border-b bg-amber-50">
        <CardTitle className="text-xl font-semibold text-amber-900">Load Out Edit</CardTitle>
      </CardHeader>
      <CardContent className="space-y-6 p-6">
        <div className="grid gap-4 md:grid-cols-3">
          <div className="space-y-2">
            <Label>Client</Label>
            <Select value={selectedClient} onValueChange={setSelectedClient}>
              <SelectTrigger>
                <SelectValue placeholder="Select client" />
              </SelectTrigger>
              <SelectContent>
                {clients.map(client => (
                  <SelectItem key={client} value={client}>
                    {client}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <div className="space-y-2">
            <Label>Work Order</Label>
            <Select
              value={selectedWorkOrderId}
              onValueChange={setSelectedWorkOrderId}
              disabled={!selectedClient}
            >
              <SelectTrigger>
                <SelectValue placeholder="Select Work Order" />
              </SelectTrigger>
              <SelectContent>
                {workOrdersForClient.map(record => (
                  <SelectItem key={`${record.id}-loadout`} value={record.wo_no}>
                    {record.wo_no}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <div className="space-y-2">
            <Label>Batch</Label>
            <Select
              value={selectedRecordId}
              onValueChange={setSelectedRecordId}
              disabled={!selectedWorkOrderId}
            >
              <SelectTrigger>
                <SelectValue placeholder="Select batch" />
              </SelectTrigger>
              <SelectContent>
                {batchesForWorkOrder.map(record => (
                  <SelectItem key={record.id} value={record.id}>
                    {record.batch} · {record.status || "Inspection Done"}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </div>

        {selectedRecordId ? (
          <div className="grid gap-4 md:grid-cols-2">
            <div className="space-y-2">
              <Label>Client</Label>
              <Input
                value={formData.client}
                onChange={event => handleInputChange("client", event.target.value)}
              />
            </div>
            <div className="space-y-2">
              <Label>Work Order</Label>
              <Input
                value={formData.wo_no}
                onChange={event => handleInputChange("wo_no", event.target.value)}
              />
            </div>
            <div className="space-y-2">
              <Label>Batch</Label>
              <Input
                value={formData.batch}
                onChange={event => handleInputChange("batch", event.target.value)}
              />
            </div>
            <div className="space-y-2">
              <Label>Status</Label>
              <Select
                value={formData.status || "Completed"}
                onValueChange={value => handleInputChange("status", value)}
              >
                <SelectTrigger>
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="Completed">Completed</SelectItem>
                  <SelectItem value="Inspection Done">Inspection Done</SelectItem>
                </SelectContent>
              </Select>
            </div>
            <div className="space-y-2">
              <Label>Load Out Date</Label>
              <DateInputField
                value={formData.load_out_date}
                onChange={value => handleInputChange("load_out_date", value)}
              />
            </div>
            <div className="space-y-2">
              <Label>AVR</Label>
              <Input
                value={formData.act_no_oper}
                onChange={event => handleInputChange("act_no_oper", event.target.value)}
              />
            </div>
            <div className="space-y-2">
              <Label>AVR Date</Label>
              <DateInputField
                value={formData.act_date}
                onChange={value => handleInputChange("act_date", value)}
              />
            </div>
          </div>
        ) : (
          <div className="rounded-lg border border-dashed border-amber-300 bg-amber-50 p-6 text-center text-sm text-amber-900">
            Select a batch with status Completed or Inspection Done to edit load out data.
          </div>
        )}

        <div className="flex justify-end">
          <Button onClick={handleUpdate} disabled={!selectedRecordId || isSaving}>
            {isSaving ? "Updating..." : "Update Load Out"}
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}

export default function EditRecords() {
  const navigate = useNavigate();
  const { toast } = useToast();
  const { sharePointService, isConnected, refreshDataInBackground } = useSharePoint();
  const { workOrders, tubingData } = useSharePointInstantData();

  const workOrderRecords = useMemo(() => parseWorkOrders(workOrders), [workOrders]);
  const tubingRecords = useMemo(() => parseTubingRecords(tubingData), [tubingData]);

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
            <ArrowLeft className="h-4 w-4" />
            <span>Back to Dashboard</span>
          </Button>
        </div>

        <Card className="mb-6 border-2 shadow-sm">
          <CardHeader className="border-b bg-white">
            <CardTitle className="text-xl font-semibold">Edit Records</CardTitle>
          </CardHeader>
          <CardContent className="space-y-2 text-sm text-slate-600">
            <p>Исправляйте данные во всех ключевых регистрах без потери автоматических расчётов.</p>
            <ul className="list-disc space-y-1 pl-5">
              <li>Перед обновлением убедитесь, что подключены к SharePoint.</li>
              <li>Автоматические поля (Key Column, Pipe From/To, этапы инспекции) пересчитываются автоматически.</li>
              <li>После сохранения данные будут обновлены в общей таблице и попадут в другие карточки.</li>
            </ul>
          </CardContent>
        </Card>

        <Tabs defaultValue="work-order" className="space-y-6">
          <TabsList className="grid w-full grid-cols-4 border-2">
            <TabsTrigger value="work-order" className="font-semibold">Work Order</TabsTrigger>
            <TabsTrigger value="tubing" className="font-semibold">Tubing Registry</TabsTrigger>
            <TabsTrigger value="inspection" className="font-semibold">Inspection</TabsTrigger>
            <TabsTrigger value="loadout" className="font-semibold">Load Out</TabsTrigger>
          </TabsList>

          <TabsContent value="work-order" className="space-y-6">
            <WorkOrderEditSection
              records={workOrderRecords}
              sharePointService={sharePointService}
              isConnected={isConnected}
              refreshData={refreshDataInBackground}
              toast={toast}
            />
          </TabsContent>

          <TabsContent value="tubing" className="space-y-6">
            <TubingEditSection
              records={tubingRecords}
              sharePointService={sharePointService}
              isConnected={isConnected}
              refreshData={refreshDataInBackground}
              toast={toast}
            />
          </TabsContent>

          <TabsContent value="inspection" className="space-y-6">
            <InspectionEditSection
              records={tubingRecords}
              sharePointService={sharePointService}
              isConnected={isConnected}
              refreshData={refreshDataInBackground}
              toast={toast}
            />
          </TabsContent>

          <TabsContent value="loadout" className="space-y-6">
            <LoadOutEditSection
              records={tubingRecords}
              sharePointService={sharePointService}
              isConnected={isConnected}
              refreshData={refreshDataInBackground}
              toast={toast}
            />
          </TabsContent>
        </Tabs>
      </div>
    </div>
  );
}

