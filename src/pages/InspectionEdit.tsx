import { useEffect, useMemo, useRef, useState } from "react";
import { useLocation, useNavigate } from "react-router-dom";
import { ArrowLeft, ClipboardCheck } from "lucide-react";

import { Header } from "@/components/layout/Header";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { DateInputField } from "@/components/ui/date-input";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { useToast } from "@/hooks/use-toast";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useSharePointInstantData } from "@/hooks/useInstantData";
import {
  parseTubingRecords,
  sanitizeNumberString,
  StageKey,
  ScrapKey
} from "@/lib/tubing-records";

interface LocationState {
  client?: string;
  wo_no?: string;
  batch?: string;
}

const STAGE_META: {
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

export default function InspectionEdit() {
  const navigate = useNavigate();
  const location = useLocation();
  const { toast } = useToast();
  const { sharePointService, isConnected, refreshDataInBackground } = useSharePoint();
  const { tubingData } = useSharePointInstantData();

  const { client, wo_no, batch } = (location.state as LocationState | null) ?? {};

  const records = useMemo(() => parseTubingRecords(tubingData ?? []), [tubingData]);
  const record = useMemo(
    () =>
      records.find(
        item =>
          item.client === client &&
          item.wo_no === wo_no &&
          item.batch === batch
      ) ?? null,
    [records, client, wo_no, batch]
  );

  const [baseStageQuantities, setBaseStageQuantities] = useState<Record<StageKey, string>>({
    rattling: "",
    external: "",
    hydro: "",
    mpi: "",
    drift: "",
    emi: "",
    marking: ""
  });
  const [scrapQuantities, setScrapQuantities] = useState<Record<ScrapKey, string>>({
    rattling: "",
    external: "",
    jetting: "",
    mpi: "",
    drift: "",
    emi: ""
  });
  const [class1, setClass1] = useState("");
  const [class2, setClass2] = useState("");
  const [class3, setClass3] = useState("");
  const [repair, setRepair] = useState("");
  const [scrapValue, setScrapValue] = useState("");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [isSaving, setIsSaving] = useState(false);
  const lastRecordIdRef = useRef<string | null>(null);
  const initialRef = useRef<{
    base: Record<StageKey, string>;
    scrap: Record<ScrapKey, string>;
    class1: string;
    class2: string;
    class3: string;
    repair: string;
    startDate: string;
    endDate: string;
  } | null>(null);

  useEffect(() => {
    if (!record) {
      lastRecordIdRef.current = null;
      return;
    }

    if (lastRecordIdRef.current === record.id) {
      return;
    }

    lastRecordIdRef.current = record.id;

    const fallbackQty = record.qty || "";

    setBaseStageQuantities({
      rattling: record.quantities.rattling ?? fallbackQty,
      external: record.quantities.external ?? fallbackQty,
      hydro: record.quantities.hydro ?? fallbackQty,
      mpi: record.quantities.mpi ?? fallbackQty,
      drift: record.quantities.drift ?? fallbackQty,
      emi: record.quantities.emi ?? fallbackQty,
      marking: record.quantities.marking ?? fallbackQty
    });

    setScrapQuantities({
      rattling: record.scrap.rattling ?? "",
      external: record.scrap.external ?? "",
      jetting: record.scrap.jetting ?? "",
      mpi: record.scrap.mpi ?? "",
      drift: record.scrap.drift ?? "",
      emi: record.scrap.emi ?? ""
    });

    setClass1(record.class_1 || "");
    setClass2(record.class_2 || "");
    setClass3(record.class_3 || "");
    setRepair(record.repair || "");
    setScrapValue(record.scrapTotal || "");
    setStartDate(record.start_date || "");
    setEndDate(record.end_date || "");

    initialRef.current = {
      base: {
        rattling: record.quantities.rattling ?? fallbackQty,
        external: record.quantities.external ?? fallbackQty,
        hydro: record.quantities.hydro ?? fallbackQty,
        mpi: record.quantities.mpi ?? fallbackQty,
        drift: record.quantities.drift ?? fallbackQty,
        emi: record.quantities.emi ?? fallbackQty,
        marking: record.quantities.marking ?? fallbackQty
      },
      scrap: {
        rattling: record.scrap.rattling ?? "",
        external: record.scrap.external ?? "",
        jetting: record.scrap.jetting ?? "",
        mpi: record.scrap.mpi ?? "",
        drift: record.scrap.drift ?? "",
        emi: record.scrap.emi ?? ""
      },
      class1: record.class_1 || "",
      class2: record.class_2 || "",
      class3: record.class_3 || "",
      repair: record.repair || "",
      startDate: record.start_date || "",
      endDate: record.end_date || ""
    };
  }, [record]);

  const handleScrapChange = (key: ScrapKey, value: string) => {
    const sanitized = sanitizeNumberString(value);
    setScrapQuantities(prev => ({ ...prev, [key]: sanitized }));
  };

  const isDirty = useMemo(() => {
    if (!initialRef.current) return false;
    const init = initialRef.current;
    const baseChanged = Object.keys(baseStageQuantities).some(
      key => (baseStageQuantities as any)[key] !== (init.base as any)[key]
    );
    const scrapChanged = Object.keys(scrapQuantities).some(
      key => (scrapQuantities as any)[key] !== (init.scrap as any)[key]
    );
    return (
      baseChanged ||
      scrapChanged ||
      class1 !== init.class1 ||
      class2 !== init.class2 ||
      class3 !== init.class3 ||
      repair !== init.repair ||
      startDate !== init.startDate ||
      endDate !== init.endDate
    );
  }, [baseStageQuantities, scrapQuantities, class1, class2, class3, repair, startDate, endDate]);

  // warn on closing tab/browser with unsaved changes
  useEffect(() => {
    const handler = (e: BeforeUnloadEvent) => {
      if (isDirty) {
        e.preventDefault();
        e.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", handler);
    return () => window.removeEventListener("beforeunload", handler);
  }, [isDirty]);

  const handleBack = () => {
    if (isDirty && !confirm("Discard your changes? Changes will not be saved.")) {
      return;
    }
    navigate("/edit-records");
  };

  const handleCancel = () => {
    if (!initialRef.current) {
      navigate("/");
      return;
    }
    const init = initialRef.current;
    setBaseStageQuantities(init.base);
    setScrapQuantities(init.scrap);
    setClass1(init.class1);
    setClass2(init.class2);
    setClass3(init.class3);
    setRepair(init.repair);
    setStartDate(init.startDate);
    setEndDate(init.endDate);
    navigate("/");
  };

  const computedScrapTotal = useMemo(
    () =>
      (Object.values(scrapQuantities) as string[]).reduce((acc, current) => {
        const numeric = Number(sanitizeNumberString(current));
        return acc + (Number.isFinite(numeric) ? numeric : 0);
      }, 0),
    [scrapQuantities]
  );

  useEffect(() => {
    const hasScrapEntries = (Object.values(scrapQuantities) as string[]).some(value => sanitizeNumberString(value) !== "");
    const computedString = hasScrapEntries ? computedScrapTotal.toString() : "";
    if (scrapValue !== computedString) {
      setScrapValue(computedString);
    }
  }, [computedScrapTotal, scrapQuantities, scrapValue]);

  const calculatedStageQuantities = useMemo(() => {
    const result: Record<StageKey, string> = {
      rattling: "",
      external: "",
      hydro: "",
      mpi: "",
      drift: "",
      emi: "",
      marking: ""
    };

    for (const stage of STAGE_META) {
      const baseRaw = baseStageQuantities[stage.key] ?? "";
      const baseNumeric = Number(sanitizeNumberString(baseRaw));
      const hasValidBase = baseRaw !== "" && Number.isFinite(baseNumeric);

      if (stage.scrapKey) {
        const scrapRaw = scrapQuantities[stage.scrapKey] ?? "";
        const scrapNumeric = Number(sanitizeNumberString(scrapRaw));
        if (hasValidBase) {
          const computed = Math.max(0, baseNumeric - (Number.isFinite(scrapNumeric) ? scrapNumeric : 0));
          result[stage.key] = computed.toString();
        } else {
          result[stage.key] = baseRaw;
        }
      } else {
        result[stage.key] = baseRaw;
      }
    }

    return result;
  }, [baseStageQuantities, scrapQuantities]);


  const handleUpdate = async () => {
    if (!record) {
      toast({
        title: "Batch not found",
        description: "Return to Edit Records to choose a batch.",
        variant: "destructive"
      });
      return;
    }

    if (!sharePointService || !isConnected) {
      toast({
        title: "SharePoint not connected",
        description: "Connect to SharePoint before updating records.",
        variant: "destructive"
      });
      return;
    }

    setIsSaving(true);
    try {
      const success = await sharePointService.updateTubingInspectionData({
        client: record.client,
        wo_no: record.wo_no,
        batch: record.batch,
        class_1: class1,
        class_2: class2,
        class_3: class3,
        repair,
        start_date: startDate,
        end_date: endDate,
        rattling_qty: Number(calculatedStageQuantities.rattling || 0),
        external_qty: Number(calculatedStageQuantities.external || 0),
        hydro_qty: Number(calculatedStageQuantities.hydro || 0),
        mpi_qty: Number(calculatedStageQuantities.mpi || 0),
        drift_qty: Number(calculatedStageQuantities.drift || 0),
        emi_qty: Number(calculatedStageQuantities.emi || 0),
        marking_qty: Number(calculatedStageQuantities.marking || 0),
        originalClient: record.originalClient,
        originalWo: record.originalWo,
        originalBatch: record.originalBatch
      });

      if (!success) {
        toast({
          title: "Update failed",
          description: "Unable to update inspection data. Please try again.",
          variant: "destructive"
        });
        return;
      }

      toast({ title: "Changes saved", description: `${record.batch} updated successfully.` });

      import("@/lib/safe-storage").then(({ safeLocalStorage }) => {
        try { safeLocalStorage.removeItem("sharepoint_last_refresh"); } catch {}
      });
      await refreshDataInBackground(sharePointService);
      navigate("/edit-records");
    } catch (error) {
      console.error("Failed to update inspection data", error);
      toast({
        title: "Update failed",
        description: "Unexpected error occurred while saving inspection data.",
        variant: "destructive"
      });
    } finally {
      setIsSaving(false);
    }
  };

  const missingSelection = !client || !wo_no || !batch;

  return (
    <div className="min-h-screen bg-slate-50">
      <Header />
      <main className="container mx-auto px-4 py-6 space-y-6">
        <div className="flex items-center justify-between">
          <Button variant="ghost" onClick={handleBack} className="flex items-center gap-2 text-slate-600">
            <ArrowLeft className="h-4 w-4" />
            Back to Edit Records
          </Button>
          <div className="flex items-center gap-2 text-sm text-blue-600">
            <ClipboardCheck className="h-4 w-4" />
            Inspection Edit
          </div>
        </div>

        {missingSelection || !record ? (
          <Card className="border-2 border-dashed border-blue-200 bg-white">
            <CardContent className="p-6 text-center text-sm text-blue-700">
              Batch details not found. Please return to Edit Records and select a batch.
            </CardContent>
          </Card>
        ) : (
          <div className="grid gap-5 lg:grid-cols-[1.1fr,1fr]">
            <Card className="border-2 border-blue-200 shadow-sm">
              <CardHeader className="border-b bg-white/80">
                <CardTitle className="text-xl font-semibold text-blue-900">Inspection Stages</CardTitle>
              </CardHeader>
              <CardContent className="space-y-4 p-5">
                <div className="grid gap-3 rounded-xl border border-blue-100 bg-blue-50/70 p-3 md:grid-cols-4">
                  <div>
                    <p className="text-xs uppercase tracking-wide text-blue-700">Client</p>
                    <p className="text-base font-semibold text-blue-900">{record.client}</p>
                  </div>
                  <div>
                    <p className="text-xs uppercase tracking-wide text-blue-700">Work Order</p>
                    <p className="text-base font-semibold text-blue-900">{record.wo_no}</p>
                  </div>
                  <div>
                    <p className="text-xs uppercase tracking-wide text-blue-700">Batch</p>
                    <p className="text-base font-semibold text-blue-900">{record.batch}</p>
                  </div>
                  <div>
                    <p className="text-xs uppercase tracking-wide text-blue-700">Quantity</p>
                    <p className="text-base font-semibold text-blue-900">{record.qty || "0"}</p>
                  </div>
                </div>

                <div className="overflow-x-auto rounded-lg border border-blue-100">
                  <Table>
                    <TableHeader className="bg-blue-50 [&_th]:h-10 [&_th]:px-3">
                      <TableRow>
                        <TableHead className="w-1/3 text-sm font-semibold text-blue-700">Stage</TableHead>
                        <TableHead className="text-sm font-semibold text-blue-700">Qty</TableHead>
                        <TableHead className="text-sm font-semibold text-blue-700">Scrap Qty</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {STAGE_META.map(stage => (
                        <TableRow key={stage.key}>
                          <TableCell className="p-3 text-sm font-medium text-slate-700">{stage.label}</TableCell>
                          <TableCell className="p-3">
                            <Input
                              value={calculatedStageQuantities[stage.key] ?? ""}
                              readOnly
                              inputMode="numeric"
                              placeholder="0"
                              className="h-9 bg-slate-100"
                            />
                          </TableCell>
                          <TableCell className="p-3">
                            {stage.scrapKey ? (
                              <Input
                                value={scrapQuantities[stage.scrapKey] ?? ""}
                                onChange={event => handleScrapChange(stage.scrapKey!, event.target.value)}
                                inputMode="numeric"
                                placeholder="0"
                                className="h-9"
                              />
                            ) : (
                              <span className="text-muted-foreground">—</span>
                            )}
                          </TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>
              </CardContent>
            </Card>

            <Card className="border-2 border-emerald-200 shadow-sm">
              <CardHeader className="border-b bg-white/80">
                <CardTitle className="text-xl font-semibold text-emerald-900">Inspection Data</CardTitle>
              </CardHeader>
              <CardContent className="space-y-4 p-5">
                <div className="grid gap-3 md:grid-cols-2">
                  <div className="space-y-2">
                    <Label htmlFor="class1">Class 1</Label>
                    <Input id="class1" value={class1} onChange={event => setClass1(event.target.value)} placeholder="Enter Class 1" className="h-9" />
                  </div>
                  <div className="space-y-2">
                    <Label htmlFor="class2">Class 2</Label>
                    <Input id="class2" value={class2} onChange={event => setClass2(event.target.value)} placeholder="Enter Class 2" className="h-9" />
                  </div>
                  <div className="space-y-2">
                    <Label htmlFor="class3">Class 3</Label>
                    <Input id="class3" value={class3} onChange={event => setClass3(event.target.value)} placeholder="Enter Class 3" className="h-9" />
                  </div>
                  <div className="space-y-2">
                    <Label htmlFor="repair">Repair</Label>
                    <Input id="repair" value={repair} onChange={event => setRepair(event.target.value)} placeholder="Enter Repair" className="h-9" />
                  </div>
                </div>

                <div className="grid gap-3 md:grid-cols-2">
                  <DateInputField label="Start Date" value={startDate} onChange={setStartDate} />
                  <DateInputField label="End Date" value={endDate} onChange={setEndDate} />
                </div>

                <div className="grid gap-3 md:grid-cols-[2fr,1fr]">
                  <div className="space-y-2">
                    <Label htmlFor="scrap">Scrap</Label>
                    <Input
                      id="scrap"
                      value={scrapValue}
                      readOnly
                      placeholder="Computed automatically"
                      inputMode="numeric"
                      className="h-9 bg-slate-100"
                    />
                  </div>
                  <div className="rounded-lg border border-emerald-200 bg-emerald-50/70 p-2 text-sm text-emerald-800">
                    <p className="font-semibold">Computed Scrap</p>
                    <p>{computedScrapTotal}</p>
                  </div>
                </div>

                <div className="flex justify-end gap-3">
                  <Button variant="destructive" onClick={handleCancel} className="min-w-[120px]">Cancel</Button>
                  <Button onClick={handleUpdate} disabled={isSaving} className="min-w-[140px]">
                    {isSaving ? "Saving..." : "Save"}
                  </Button>
                </div>
              </CardContent>
            </Card>
          </div>
        )}
      </main>
    </div>
  );
}
