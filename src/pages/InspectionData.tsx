import { useEffect, useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import { Header } from "@/components/layout/Header";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { ArrowLeft, ClipboardCheck } from "lucide-react";
import { useToast } from "@/hooks/use-toast";
import { useSharePointInstantData } from "@/hooks/useInstantData";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useAuth } from "@/contexts/AuthContext";

type StageKey = "rattling" | "external" | "hydro" | "mpi" | "drift" | "emi" | "marking";

interface ArrivedBatchRow {
  key: string;
  client: string;
  wo_no: string;
  batch: string;
  status: string;
  class_1?: string;
  class_2?: string;
  class_3?: string;
  repair?: string;
  scrap?: string;
  baseQty: number | null;
  rattling_qty: number | null;
  external_qty: number | null;
  hydro_qty: number | null;
  mpi_qty: number | null;
  drift_qty: number | null;
  emi_qty: number | null;
  marking_qty: number | null;
}

const STAGE_ORDER: StageKey[] = [
  "rattling",
  "external",
  "hydro",
  "mpi",
  "drift",
  "emi",
  "marking"
];

const stageMeta: {
  key: StageKey;
  label: string;
  scrapKey?: "rattling" | "external" | "jetting" | "mpi" | "drift" | "emi";
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

const normalizeString = (value: unknown) =>
  value === null || value === undefined ? "" : String(value).trim();

const toNumeric = (value: unknown): number | null => {
  if (value === null || value === undefined) return null;
  const num = Number(String(value).replace(/[^0-9.-]/g, ""));
  return Number.isFinite(num) ? num : null;
};

const sanitizeDigits = (value: string) => value.replace(/[^0-9]/g, "");

const getPreviousStage = (stage: StageKey) => {
  const index = STAGE_ORDER.indexOf(stage);
  if (index <= 0) return null;
  return STAGE_ORDER[index - 1];
};

export default function InspectionData() {
  const navigate = useNavigate();
  const { toast } = useToast();
  const { user } = useAuth();
  const { sharePointService, isConnected } = useSharePoint();
  const { tubingData } = useSharePointInstantData();

  const [selectedClient, setSelectedClient] = useState("");
  const [selectedWorkOrder, setSelectedWorkOrder] = useState("");
  const [selectedBatch, setSelectedBatch] = useState("");
  const [selectedRow, setSelectedRow] = useState<ArrivedBatchRow | null>(null);
  const [class1, setClass1] = useState("");
  const [class2, setClass2] = useState("");
  const [class3, setClass3] = useState("");
  const [repairValue, setRepairValue] = useState("");
  const [scrapValue, setScrapValue] = useState("");
  const [stageQuantities, setStageQuantities] = useState<Record<StageKey, string>>({
    rattling: "",
    external: "",
    hydro: "",
    mpi: "",
    drift: "",
    emi: "",
    marking: ""
  });
  const [initialQty, setInitialQty] = useState<number>(0);
  const [processedKeys, setProcessedKeys] = useState<string[]>([]);
  const [isSaving, setIsSaving] = useState(false);

  const arrivedBatches = useMemo(() => {
    if (!Array.isArray(tubingData) || tubingData.length < 2) {
      return [] as ArrivedBatchRow[];
    }

    const headersRow = tubingData[0];
    if (!Array.isArray(headersRow)) {
      return [] as ArrivedBatchRow[];
    }

    const headers = headersRow as unknown[];
    const normalizeHeader = (header: unknown) =>
      header === null || header === undefined ? "" : String(header).trim().toLowerCase();

    const findIndex = (predicate: (header: string) => boolean) =>
      headers.findIndex(header => predicate(normalizeHeader(header)));

    const clientIndex = findIndex(header => header.includes("client"));
    const woIndex = findIndex(header => header.includes("wo"));
    const batchIndex = findIndex(header => header.includes("batch"));
    const statusIndex = findIndex(header => header.includes("status"));
    const baseQtyIndex = findIndex(
      header => header.includes("qty") && !header.includes("_") && !header.includes("scrap")
    );
    const class1Index = findIndex(header => header.includes("class 1") || header.includes("class_1"));
    const class2Index = findIndex(header => header.includes("class 2") || header.includes("class_2"));
    const class3Index = findIndex(header => header.includes("class 3") || header.includes("class_3"));
    const repairIndex = findIndex(header => header.includes("repair"));
    const scrapIndex = findIndex(header => header === "scrap" || header.endsWith(" scrap"));
    const rattlingQtyIndex = findIndex(header => header.includes("rattling_qty"));
    const externalQtyIndex = findIndex(header => header.includes("external_qty"));
    const hydroQtyIndex = findIndex(header => header.includes("hydro_qty"));
    const mpiQtyIndex = findIndex(header => header.includes("mpi_qty"));
    const driftQtyIndex = findIndex(header => header.includes("drift_qty"));
    const emiQtyIndex = findIndex(header => header.includes("emi_qty"));
    const markingQtyIndex = findIndex(header => header.includes("marking_qty"));

    if (statusIndex === -1 || clientIndex === -1 || woIndex === -1 || batchIndex === -1) {
      return [] as ArrivedBatchRow[];
    }

    const rows: ArrivedBatchRow[] = [];

    tubingData.slice(1).forEach(item => {
      if (!Array.isArray(item)) {
        return;
      }
      const row = item as unknown[];
      const rawStatus = normalizeString(row[statusIndex]);
      const normalizedStatus = rawStatus.toLowerCase();
      if (!normalizedStatus.includes("arriv")) {
        return;
      }

      const client = normalizeString(row[clientIndex]);
      const wo_no = normalizeString(row[woIndex]);
      const batch = normalizeString(row[batchIndex]);

      rows.push({
        key: `${client}||${wo_no}||${batch}`,
        client,
        wo_no,
        batch,
        status: rawStatus || "Arrived",
        class_1: normalizeString(class1Index === -1 ? "" : row[class1Index]),
        class_2: normalizeString(class2Index === -1 ? "" : row[class2Index]),
        class_3: normalizeString(class3Index === -1 ? "" : row[class3Index]),
        repair: normalizeString(repairIndex === -1 ? "" : row[repairIndex]),
        scrap: normalizeString(scrapIndex === -1 ? "" : row[scrapIndex]),
        baseQty: toNumeric(baseQtyIndex === -1 ? null : row[baseQtyIndex]),
        rattling_qty: toNumeric(rattlingQtyIndex === -1 ? null : row[rattlingQtyIndex]),
        external_qty: toNumeric(externalQtyIndex === -1 ? null : row[externalQtyIndex]),
        hydro_qty: toNumeric(hydroQtyIndex === -1 ? null : row[hydroQtyIndex]),
        mpi_qty: toNumeric(mpiQtyIndex === -1 ? null : row[mpiQtyIndex]),
        drift_qty: toNumeric(driftQtyIndex === -1 ? null : row[driftQtyIndex]),
        emi_qty: toNumeric(emiQtyIndex === -1 ? null : row[emiQtyIndex]),
        marking_qty: toNumeric(markingQtyIndex === -1 ? null : row[markingQtyIndex])
      });
    });

    return rows;
  }, [tubingData]);

  const availableRows = useMemo(
    () => arrivedBatches.filter(row => !processedKeys.includes(row.key)),
    [arrivedBatches, processedKeys]
  );

  const availableClients = useMemo(() => {
    const unique = new Set<string>();
    availableRows.forEach(row => {
      if (row.client) unique.add(row.client);
    });
    return Array.from(unique);
  }, [availableRows]);

  const availableWorkOrders = useMemo(() => {
    if (!selectedClient) return [] as string[];
    const unique = new Set<string>();
    availableRows
      .filter(row => row.client === selectedClient)
      .forEach(row => {
        if (row.wo_no) unique.add(row.wo_no);
      });
    return Array.from(unique);
  }, [availableRows, selectedClient]);

  const availableBatches = useMemo(() => {
    if (!selectedClient || !selectedWorkOrder) return [] as ArrivedBatchRow[];
    return availableRows.filter(
      row => row.client === selectedClient && row.wo_no === selectedWorkOrder
    );
  }, [availableRows, selectedClient, selectedWorkOrder]);

  useEffect(() => {
    if (selectedClient && !availableClients.includes(selectedClient)) {
      setSelectedClient("");
      setSelectedWorkOrder("");
      setSelectedBatch("");
      setSelectedRow(null);
    }
  }, [availableClients, selectedClient]);

  useEffect(() => {
    if (selectedWorkOrder && !availableWorkOrders.includes(selectedWorkOrder)) {
      setSelectedWorkOrder("");
      setSelectedBatch("");
      setSelectedRow(null);
    }
  }, [availableWorkOrders, selectedWorkOrder]);

  useEffect(() => {
    if (!selectedBatch) {
      setSelectedRow(null);
      return;
    }

    const match = availableBatches.find(row => row.batch === selectedBatch);
    setSelectedRow(match ?? null);
  }, [availableBatches, selectedBatch]);

  useEffect(() => {
    if (!selectedRow) {
      setClass1("");
      setClass2("");
      setClass3("");
      setRepairValue("");
      setScrapValue("");
      setStageQuantities({
        rattling: "",
        external: "",
        hydro: "",
        mpi: "",
        drift: "",
        emi: "",
        marking: ""
      });
      setInitialQty(0);
      return;
    }

    const baseQty = selectedRow.baseQty;
    const rattlingQty = selectedRow.rattling_qty;
    const base =
      baseQty !== null && baseQty !== undefined ? baseQty : rattlingQty ?? null;
    const hasBase = base != null;
    setInitialQty(hasBase ? base : 0);
    setClass1(selectedRow.class_1 || "");
    setClass2(selectedRow.class_2 || "");
    setClass3(selectedRow.class_3 || "");
    setRepairValue(selectedRow.repair || "");
    setScrapValue(selectedRow.scrap || "");
    setStageQuantities({
      rattling: hasBase ? String(base) : "",
      external:
        selectedRow.external_qty !== null && selectedRow.external_qty !== undefined
          ? String(selectedRow.external_qty)
          : "",
      hydro:
        selectedRow.hydro_qty !== null && selectedRow.hydro_qty !== undefined
          ? String(selectedRow.hydro_qty)
          : "",
      mpi:
        selectedRow.mpi_qty !== null && selectedRow.mpi_qty !== undefined
          ? String(selectedRow.mpi_qty)
          : "",
      drift:
        selectedRow.drift_qty !== null && selectedRow.drift_qty !== undefined
          ? String(selectedRow.drift_qty)
          : "",
      emi:
        selectedRow.emi_qty !== null && selectedRow.emi_qty !== undefined
          ? String(selectedRow.emi_qty)
          : "",
      marking:
        selectedRow.marking_qty !== null && selectedRow.marking_qty !== undefined
          ? String(selectedRow.marking_qty)
          : ""
    });
  }, [selectedRow]);

  const scrapValues = useMemo(() => {
    const parse = (value: string) => {
      if (value === "") return null;
      const num = Number(value);
      return Number.isFinite(num) ? num : null;
    };

    const rattling = parse(stageQuantities.rattling);
    const external = parse(stageQuantities.external);
    const hydro = parse(stageQuantities.hydro);
    const mpi = parse(stageQuantities.mpi);
    const drift = parse(stageQuantities.drift);
    const emi = parse(stageQuantities.emi);
    const marking = parse(stageQuantities.marking);

    const diff = (prev: number | null, next: number | null) => {
      if (prev === null || next === null) return null;
      if (next > prev) return null;
      return prev - next;
    };

    return {
      rattling: diff(rattling, external),
      external: diff(external, hydro),
      jetting: diff(hydro, mpi),
      mpi: diff(mpi, drift),
      drift: diff(drift, emi),
      emi: diff(emi, marking)
    };
  }, [stageQuantities]);

  const totalScrap = useMemo(
    () =>
      Object.values(scrapValues).reduce((sum, value) => (value !== null ? sum + value : sum), 0),
    [scrapValues]
  );

  const handleQuantityChange = (
    stage: StageKey,
    value: string,
    options?: { allowManualRattling?: boolean }
  ) => {
    const sanitized = sanitizeDigits(value);
    if (sanitized === "") {
      setStageQuantities(prev => ({ ...prev, [stage]: "" }));
      if (stage === "rattling") {
        setInitialQty(0);
      }
      return;
    }

    const numericValue = Number(sanitized);
    if (!Number.isFinite(numericValue)) {
      return;
    }

    if (stage === "rattling") {
      if (options?.allowManualRattling) {
        setStageQuantities(prev => ({ ...prev, rattling: sanitized }));
        setInitialQty(numericValue);
      }
      return;
    }

    const previousStage = getPreviousStage(stage);
    if (previousStage) {
      const previousValue = Number(stageQuantities[previousStage]);
      if (stageQuantities[previousStage] !== "" && numericValue > previousValue) {
        toast({
          title: "Ошибка",
          description: "Количество на следующем этапе не может превышать предыдущее",
          variant: "destructive"
        });
        return;
      }
    }

    setStageQuantities(prev => ({ ...prev, [stage]: sanitized }));
  };

  const handleSave = async () => {
    if (!user) {
      toast({
        title: "Ошибка",
        description: "Пожалуйста, войдите в систему",
        variant: "destructive"
      });
      return;
    }

    if (!sharePointService || !isConnected) {
      toast({
        title: "Ошибка",
        description: "SharePoint не подключен",
        variant: "destructive"
      });
      return;
    }

    if (!selectedRow) {
      toast({
        title: "Ошибка",
        description: "Выберите партию для сохранения",
        variant: "destructive"
      });
      return;
    }

    const stageNumbers: Record<StageKey, number> = {
      rattling: Number(stageQuantities.rattling),
      external: Number(stageQuantities.external),
      hydro: Number(stageQuantities.hydro),
      mpi: Number(stageQuantities.mpi),
      drift: Number(stageQuantities.drift),
      emi: Number(stageQuantities.emi),
      marking: Number(stageQuantities.marking)
    };

    for (const stage of STAGE_ORDER) {
      const raw = stageQuantities[stage];
      if (raw === "" || Number.isNaN(stageNumbers[stage])) {
        toast({
          title: "Ошибка",
          description: "Заполните все количества этапов инспекции",
          variant: "destructive"
        });
        return;
      }
      if (stageNumbers[stage] < 0) {
        toast({
          title: "Ошибка",
          description: "Количество не может быть отрицательным",
          variant: "destructive"
        });
        return;
      }

      const prevStage = getPreviousStage(stage);
      if (prevStage && stageNumbers[prevStage] < stageNumbers[stage]) {
        toast({
          title: "Ошибка",
          description: "Количество на следующем этапе не может превышать предыдущее",
          variant: "destructive"
        });
        return;
      }
    }

    const canEditRattlingQty = selectedRow.rattling_qty == null && selectedRow.baseQty == null;

    if (!canEditRattlingQty && stageNumbers.rattling !== initialQty) {
      toast({
        title: "Ошибка",
        description: "Rattling Qty должно совпадать с количеством труб партии",
        variant: "destructive"
      });
      return;
    }

    const scrapInput = sanitizeDigits(scrapValue);
    if (scrapInput === "") {
      toast({
        title: "Ошибка",
        description: "Введите Scrap",
        variant: "destructive"
      });
      return;
    }

    const scrapNumber = Number(scrapInput);
    if (!Number.isFinite(scrapNumber)) {
      toast({
        title: "Ошибка",
        description: "Некорректное значение Scrap",
        variant: "destructive"
      });
      return;
    }

    const missingScrap = Object.entries(scrapValues).find(([, value]) => value === null);
    if (missingScrap) {
      toast({
        title: "Ошибка",
        description: "Проверьте таблицу — разница между этапами заполнена некорректно",
        variant: "destructive"
      });
      return;
    }

    if (scrapNumber !== totalScrap) {
      toast({
        title: "Ошибка",
        description: "Итоговый Scrap не совпадает с суммой скрапов таблицы",
        variant: "destructive"
      });
      return;
    }

    setIsSaving(true);

    const success = await sharePointService.updateTubingInspectionData({
      client: selectedRow.client,
      wo_no: selectedRow.wo_no,
      batch: selectedRow.batch,
      class_1: class1,
      class_2: class2,
      class_3: class3,
      repair: sanitizeDigits(repairValue) || "0",
      scrap: scrapNumber,
      rattling_qty: stageNumbers.rattling,
      external_qty: stageNumbers.external,
      hydro_qty: stageNumbers.hydro,
      mpi_qty: stageNumbers.mpi,
      drift_qty: stageNumbers.drift,
      emi_qty: stageNumbers.emi,
      marking_qty: stageNumbers.marking,
      rattling_scrap_qty: scrapValues.rattling ?? 0,
      external_scrap_qty: scrapValues.external ?? 0,
      jetting_scrap_qty: scrapValues.jetting ?? 0,
      mpi_scrap_qty: scrapValues.mpi ?? 0,
      drift_scrap_qty: scrapValues.drift ?? 0,
      emi_scrap_qty: scrapValues.emi ?? 0,
      status: "Inspection Done"
    });

    setIsSaving(false);

    if (success) {
      toast({
        title: "Успешно",
        description: "Инспекция сохранена и партия обновлена",
        variant: "default"
      });
      setProcessedKeys(prev => (prev.includes(selectedRow.key) ? prev : [...prev, selectedRow.key]));
      setSelectedBatch("");
    } else {
      toast({
        title: "Ошибка",
        description: "Не удалось обновить данные партии",
        variant: "destructive"
      });
    }
  };

  return (
    <div className="min-h-screen bg-gray-50">
      <Header />
      <div className="container mx-auto px-6 py-8">
        <div className="mb-6 flex flex-wrap items-center justify-between gap-4">
          <Button variant="outline" onClick={() => navigate("/")} className="flex items-center gap-2">
            <ArrowLeft className="h-4 w-4" />
            <span>Back to Dashboard</span>
          </Button>
          <div className="flex items-center gap-2 text-gray-600">
            <ClipboardCheck className="h-5 w-5" />
            <span>Inspection Data Entry</span>
          </div>
        </div>

        <div className="grid gap-6 lg:grid-cols-5">
          <Card className="lg:col-span-2">
            <CardHeader>
              <CardTitle className="text-xl font-semibold text-blue-900">Batch Selection</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="space-y-2">
                <Label>Client</Label>
                <Select
                  value={selectedClient || undefined}
                  onValueChange={value => setSelectedClient(value)}
                >
                  <SelectTrigger>
                    <SelectValue placeholder="Choose client" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableClients.length === 0 && (
                      <div className="px-2 py-1 text-sm text-muted-foreground">No arrived batches</div>
                    )}
                    {availableClients.map(client => (
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
                  value={selectedWorkOrder || undefined}
                  onValueChange={value => setSelectedWorkOrder(value)}
                  disabled={!selectedClient}
                >
                  <SelectTrigger>
                    <SelectValue placeholder="Choose work order" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableWorkOrders.length === 0 && (
                      <div className="px-2 py-1 text-sm text-muted-foreground">No arrived batches</div>
                    )}
                    {availableWorkOrders.map(wo => (
                      <SelectItem key={wo} value={wo}>
                        {wo}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-2">
                <Label>Batch</Label>
                <Select
                  value={selectedBatch || undefined}
                  onValueChange={value => setSelectedBatch(value)}
                  disabled={!selectedClient || !selectedWorkOrder}
                >
                  <SelectTrigger>
                    <SelectValue placeholder="Choose arrived batch" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableBatches.length === 0 && (
                      <div className="px-2 py-1 text-sm text-muted-foreground">No arrived batches</div>
                    )}
                    {availableBatches.map(batch => (
                      <SelectItem key={batch.batch} value={batch.batch}>
                        {batch.batch}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              {selectedRow && (
                <div className="rounded-lg bg-blue-50 p-4 text-sm text-blue-900">
                  <p className="font-semibold">Batch Info</p>
                  <p>Qty: {initialQty}</p>
                  <p>Status: {selectedRow.status || "Arrived"}</p>
                </div>
              )}
            </CardContent>
          </Card>

          <Card className="lg:col-span-3">
            <CardHeader>
              <CardTitle className="text-xl font-semibold text-emerald-900">Inspection Data</CardTitle>
            </CardHeader>
            <CardContent className="space-y-6">
              <div className="grid gap-4 md:grid-cols-2">
                <div className="space-y-2">
                  <Label htmlFor="class1">Class 1</Label>
                  <Input
                    id="class1"
                    value={class1}
                    onChange={event => setClass1(event.target.value)}
                    placeholder="Enter Class 1"
                  />
                </div>
                <div className="space-y-2">
                  <Label htmlFor="class2">Class 2</Label>
                  <Input
                    id="class2"
                    value={class2}
                    onChange={event => setClass2(event.target.value)}
                    placeholder="Enter Class 2"
                  />
                </div>
                <div className="space-y-2">
                  <Label htmlFor="class3">Class 3</Label>
                  <Input
                    id="class3"
                    value={class3}
                    onChange={event => setClass3(event.target.value)}
                    placeholder="Enter Class 3"
                  />
                </div>
                <div className="space-y-2">
                  <Label htmlFor="repair">Repair</Label>
                  <Input
                    id="repair"
                    value={repairValue}
                    onChange={event => setRepairValue(sanitizeDigits(event.target.value))}
                    placeholder="0"
                    inputMode="numeric"
                  />
                </div>
                <div className="space-y-2 md:col-span-2">
                  <Label htmlFor="scrap">Scrap</Label>
                  <Input
                    id="scrap"
                    value={scrapValue}
                    onChange={event => setScrapValue(sanitizeDigits(event.target.value))}
                    placeholder="0"
                    inputMode="numeric"
                  />
                </div>
              </div>

              <div className="space-y-4">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead className="w-1/3">Stage</TableHead>
                      <TableHead>Qty</TableHead>
                      <TableHead>Scrap Qty</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {stageMeta.map(stage => (
                      <TableRow key={stage.key}>
                        <TableCell className="font-medium">{stage.label}</TableCell>
                        <TableCell>
                          <Input
                            value={stageQuantities[stage.key]}
                            onChange={event =>
                              handleQuantityChange(stage.key, event.target.value, {
                                allowManualRattling:
                                  stage.key === "rattling" &&
                                  selectedRow != null &&
                                  selectedRow.rattling_qty == null &&
                                  selectedRow.baseQty == null
                              })
                            }
                            inputMode="numeric"
                            disabled={
                              !selectedRow ||
                              (stage.key === "rattling" &&
                                (selectedRow.rattling_qty != null || selectedRow.baseQty != null))
                            }
                            placeholder="0"
                          />
                        </TableCell>
                        <TableCell>
                          {stage.scrapKey ? (
                            <span className="font-mono text-sm">
                              {scrapValues[stage.scrapKey] !== null ? scrapValues[stage.scrapKey] : "—"}
                            </span>
                          ) : (
                            <span className="text-muted-foreground">—</span>
                          )}
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>

              <div className="flex flex-wrap items-center justify-between gap-4">
                <div className="text-sm text-muted-foreground">
                  Итоговый Scrap: <span className="font-semibold text-emerald-700">{totalScrap}</span>
                </div>
                <Button onClick={handleSave} disabled={isSaving || !selectedRow}>
                  {isSaving ? "Saving..." : "Save"}
                </Button>
              </div>
            </CardContent>
          </Card>
        </div>
      </div>
    </div>
  );
}
