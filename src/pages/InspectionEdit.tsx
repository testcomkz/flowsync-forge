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
import { ConfirmDialog } from "@/components/ui/confirm-dialog";
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
  overrideQty?: number;
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

  const { client, wo_no, batch, overrideQty } = (location.state as LocationState | null) ?? {};

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

  // Effective Qty used for calculations and validations. If overrideQty provided from Batch Selection, use it.
  const effectiveQty = useMemo(() => {
    if (typeof overrideQty === 'number' && Number.isFinite(overrideQty) && overrideQty > 0) {
      return overrideQty;
    }
    const raw = sanitizeNumberString(record?.qty || "");
    const n = Number(raw);
    return Number.isFinite(n) ? n : 0;
  }, [overrideQty, record?.qty]);

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
  const [isConfirmOpen, setIsConfirmOpen] = useState(false);
  const [confirmLines, setConfirmLines] = useState<string[]>([]);
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
    // Enforce: scrap at each stage cannot exceed available quantity at that stage
    const r = effectiveQty || 0;
    const ext = Math.max(0, r - (Number(sanitizeNumberString(key === 'rattling' ? sanitized : scrapQuantities.rattling)) || 0));
    const hyd = Math.max(0, ext - (Number(sanitizeNumberString(key === 'external' ? sanitized : scrapQuantities.external)) || 0));
    const mp = Math.max(0, hyd - (Number(sanitizeNumberString(key === 'jetting' ? sanitized : scrapQuantities.jetting)) || 0));
    const dr = Math.max(0, mp - (Number(sanitizeNumberString(key === 'mpi' ? sanitized : scrapQuantities.mpi)) || 0));
    const em = Math.max(0, dr - (Number(sanitizeNumberString(key === 'drift' ? sanitized : scrapQuantities.drift)) || 0));

    const allowed: Record<ScrapKey, number> = {
      rattling: r,
      external: ext,
      jetting: hyd,
      mpi: mp,
      drift: dr,
      emi: em,
    };
    if (sanitized !== "") {
      const num = Number(sanitized);
      if (Number.isFinite(num) && num > allowed[key]) {
        toast({ title: "Ошибка", description: "Scrap не может превышать количество на текущем этапе", variant: "destructive" });
        return;
      }
    }
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
    // Recalculate from effective qty and current scraps to reflect any qty changes
    const r = effectiveQty || 0;
    const ext = Math.max(0, r - (Number(sanitizeNumberString(scrapQuantities.rattling)) || 0));
    const hyd = Math.max(0, ext - (Number(sanitizeNumberString(scrapQuantities.external)) || 0));
    const mp = Math.max(0, hyd - (Number(sanitizeNumberString(scrapQuantities.jetting)) || 0));
    const dr = Math.max(0, mp - (Number(sanitizeNumberString(scrapQuantities.mpi)) || 0));
    const em = Math.max(0, dr - (Number(sanitizeNumberString(scrapQuantities.drift)) || 0));
    const mark = Math.max(0, em - (Number(sanitizeNumberString(scrapQuantities.emi)) || 0));
    return {
      rattling: r.toString(),
      external: ext.toString(),
      hydro: hyd.toString(),
      mpi: mp.toString(),
      drift: dr.toString(),
      emi: em.toString(),
      marking: mark.toString(),
    } as Record<StageKey, string>;
  }, [effectiveQty, scrapQuantities]);


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
    // Validate classes/repair and scrap totals against batch qty
    const n = (v: string) => Number(sanitizeNumberString(v)) || 0;
    const qtyNum = effectiveQty || 0;
    const totalScrapNow = computedScrapTotal;
    const totalClasses = n(class1) + n(class2) + n(class3) + n(repair);
    if (totalClasses + totalScrapNow !== qtyNum) {
      toast({ title: "Ошибка", description: "Сумма Class1 + Class2 + Class3 + Repair + Scrap должна равняться Qty батча", variant: "destructive" });
      return;
    }
    // Per-stage scrap must not exceed available quantities (handles qty changes)
    const allowedMap: Record<ScrapKey, number> = {
      rattling: Number(calculatedStageQuantities.rattling) || 0,
      external: Number(calculatedStageQuantities.external) || 0,
      jetting: Number(calculatedStageQuantities.hydro) || 0,
      mpi: Number(calculatedStageQuantities.mpi) || 0,
      drift: Number(calculatedStageQuantities.drift) || 0,
      emi: Number(calculatedStageQuantities.emi) || 0,
    };
    for (const key of Object.keys(scrapQuantities) as ScrapKey[]) {
      const val = n(scrapQuantities[key]);
      if (val > allowedMap[key]) {
        toast({ title: "Ошибка", description: "Scrap Qty exceeds new Qty. Change it to continue.", variant: "destructive" });
        return;
      }
    }
    setConfirmLines([
      `Client: ${record.client}`,
      `WO: ${record.wo_no}`,
      `Batch: ${record.batch}`
    ]);
    setIsConfirmOpen(true);
  };

  const doUpdate = async () => {
    if (!record || !sharePointService) return;
    setIsSaving(true);
    try {
      const num = (v: string) => Number(sanitizeNumberString(v)) || 0;
      const success = await sharePointService.updateTubingInspectionData({
        client: record.client,
        wo_no: record.wo_no,
        batch: record.batch,
        class_1: String(num(class1)),
        class_2: String(num(class2)),
        class_3: String(num(class3)),
        repair: String(num(repair)),
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
      const statusNorm = (record.status || "").toLowerCase();
      if (statusNorm.includes("completed")) {
        navigate("/load-out-edit", { state: { client: record.client, wo_no: record.wo_no, batch: record.batch } });
      } else {
        navigate("/edit-records");
      }
    } catch (error) {
      console.error("Failed to update inspection data", error);
      toast({
        title: "Update failed",
        description: "Unexpected error occurred while saving inspection data.",
        variant: "destructive"
      });
    } finally {
      setIsSaving(false);
      setIsConfirmOpen(false);
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

        <ConfirmDialog
          open={isConfirmOpen}
          title="Update Inspection?"
          description="Confirm saving changes for this batch"
          lines={confirmLines}
          confirmText="Save"
          cancelText="Cancel"
          onConfirm={doUpdate}
          onCancel={() => setIsConfirmOpen(false)}
          loading={isSaving}
        />
        {missingSelection || !record ? (
          <Card className="border-2 border-dashed border-blue-200 bg-white">
            <CardContent className="p-6 text-center text-sm text-blue-700">
              Batch details not found. Please return to Edit Records and select a batch.
            </CardContent>
          </Card>
        ) : (
          <div className="grid gap-5 lg:grid-cols-[1.1fr,1fr]">
            <Card className="border-2 border-blue-200 rounded-xl shadow-md">
              <CardHeader className="border-b bg-blue-50">
                <CardTitle className="text-xl font-semibold text-blue-900">Inspection Stages</CardTitle>
              </CardHeader>
              <CardContent className="space-y-4 p-5">
                <div className="grid gap-3 rounded-xl border border-blue-100 bg-white p-3 md:grid-cols-4">
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
                    <p className="text-base font-semibold text-blue-900">{String(effectiveQty)}</p>
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

            <Card className="border-2 border-blue-200 rounded-xl shadow-md">
              <CardHeader className="border-b bg-blue-50">
                <CardTitle className="text-xl font-semibold text-blue-900">Inspection Data</CardTitle>
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

                <div className="grid gap-3 md:grid-cols-1">
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
                </div>

                <div className="flex justify-end gap-3">
                  <Button variant="destructive" onClick={handleCancel} className="min-w-[120px]">Cancel</Button>
                  <Button onClick={handleUpdate} disabled={isSaving} className="min-w-[140px] bg-blue-600 hover:bg-blue-700 text-white">
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
