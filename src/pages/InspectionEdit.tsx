import { useEffect, useMemo, useState } from "react";
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
import { useUnsavedChangesWarning } from "@/hooks/use-unsaved-changes";

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

  const defaultInitial = useMemo(
    () => ({
      scrapQuantities: {
        rattling: "",
        external: "",
        jetting: "",
        mpi: "",
        drift: "",
        emi: "",
      } satisfies Record<ScrapKey, string>,
      class1: "",
      class2: "",
      class3: "",
      repair: "",
      scrapValue: "",
      startDate: "",
      endDate: "",
    }),
    []
  );

  const [initialValues, setInitialValues] = useState(defaultInitial);
  const [scrapQuantities, setScrapQuantities] = useState<Record<ScrapKey, string>>(defaultInitial.scrapQuantities);
  const [class1, setClass1] = useState(defaultInitial.class1);
  const [class2, setClass2] = useState(defaultInitial.class2);
  const [class3, setClass3] = useState(defaultInitial.class3);
  const [repair, setRepair] = useState(defaultInitial.repair);
  const [scrapValue, setScrapValue] = useState(defaultInitial.scrapValue);
  const [startDate, setStartDate] = useState(defaultInitial.startDate);
  const [endDate, setEndDate] = useState(defaultInitial.endDate);
  const [isSaving, setIsSaving] = useState(false);

  useEffect(() => {
    if (!record) {
      setInitialValues(defaultInitial);
      setScrapQuantities(defaultInitial.scrapQuantities);
      setClass1(defaultInitial.class1);
      setClass2(defaultInitial.class2);
      setClass3(defaultInitial.class3);
      setRepair(defaultInitial.repair);
      setScrapValue(defaultInitial.scrapValue);
      setStartDate(defaultInitial.startDate);
      setEndDate(defaultInitial.endDate);
      return;
    }

    const nextInitial = {
      scrapQuantities: {
        rattling: sanitizeNumberString(record.scrap.rattling ?? ""),
        external: sanitizeNumberString(record.scrap.external ?? ""),
        jetting: sanitizeNumberString(record.scrap.jetting ?? ""),
        mpi: sanitizeNumberString(record.scrap.mpi ?? ""),
        drift: sanitizeNumberString(record.scrap.drift ?? ""),
        emi: sanitizeNumberString(record.scrap.emi ?? ""),
      } as Record<ScrapKey, string>,
      class1: record.class_1 || "",
      class2: record.class_2 || "",
      class3: record.class_3 || "",
      repair: record.repair || "",
      scrapValue: record.scrapTotal || "",
      startDate: record.start_date || "",
      endDate: record.end_date || "",
    };

    setInitialValues({ ...nextInitial, scrapQuantities: { ...nextInitial.scrapQuantities } });
    setScrapQuantities({ ...nextInitial.scrapQuantities });
    setClass1(nextInitial.class1);
    setClass2(nextInitial.class2);
    setClass3(nextInitial.class3);
    setRepair(nextInitial.repair);
    setScrapValue(nextInitial.scrapValue);
    setStartDate(nextInitial.startDate);
    setEndDate(nextInitial.endDate);
  }, [defaultInitial, record]);

  const handleScrapChange = (key: ScrapKey, value: string) => {
    const sanitized = sanitizeNumberString(value);
    setScrapQuantities(prev => ({ ...prev, [key]: sanitized }));
  };

  const handleBack = () => {
    if (record) {
      navigate("/edit-records", {
        state: { client: record.client, wo_no: record.wo_no, batch: record.batch },
      });
      return;
    }
    navigate("/edit-records");
  };

  const computedScrapTotal = useMemo(
    () =>
      (Object.values(scrapQuantities) as string[]).reduce((acc, current) => {
        const numeric = Number(sanitizeNumberString(current));
        return acc + (Number.isFinite(numeric) ? numeric : 0);
      }, 0),
    [scrapQuantities]
  );

  const computedStageQuantities = useMemo(() => {
    if (!record) {
      return {
        rattling: "",
        external: "",
        hydro: "",
        mpi: "",
        drift: "",
        emi: "",
        marking: "",
      } satisfies Record<StageKey, string>;
    }

    const baseCandidates = [
      record.qty,
      record.quantities.rattling,
      record.quantities.external,
      record.quantities.hydro,
      record.quantities.mpi,
      record.quantities.drift,
      record.quantities.emi,
      record.quantities.marking,
    ];

    const initialValue = baseCandidates
      .map(value => sanitizeNumberString(value ?? ""))
      .find(value => value !== "");

    const hasInitialValue = Boolean(initialValue);
    let running = hasInitialValue ? Number(initialValue) : 0;
    if (!Number.isFinite(running)) {
      running = 0;
    }

    const result: Record<StageKey, string> = {
      rattling: "",
      external: "",
      hydro: "",
      mpi: "",
      drift: "",
      emi: "",
      marking: "",
    };

    STAGE_META.forEach(stage => {
      result[stage.key] = hasInitialValue ? String(Math.max(0, Math.trunc(running))) : "";

      if (stage.scrapKey) {
        const sanitizedScrap = sanitizeNumberString(scrapQuantities[stage.scrapKey] ?? "");
        const scrapValue = sanitizedScrap ? Number(sanitizedScrap) : 0;
        if (Number.isFinite(scrapValue)) {
          running = Math.max(0, running - scrapValue);
        }
      }
    });

    return result;
  }, [record, scrapQuantities]);

  const toNumericValue = (value: string) => {
    const sanitized = sanitizeNumberString(value);
    if (!sanitized) {
      return 0;
    }
    const numeric = Number(sanitized);
    return Number.isFinite(numeric) ? numeric : 0;
  };

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
        scrap: scrapValue,
        start_date: startDate,
        end_date: endDate,
        rattling_qty: toNumericValue(computedStageQuantities.rattling ?? ""),
        external_qty: toNumericValue(computedStageQuantities.external ?? ""),
        hydro_qty: toNumericValue(computedStageQuantities.hydro ?? ""),
        mpi_qty: toNumericValue(computedStageQuantities.mpi ?? ""),
        drift_qty: toNumericValue(computedStageQuantities.drift ?? ""),
        emi_qty: toNumericValue(computedStageQuantities.emi ?? ""),
        marking_qty: toNumericValue(computedStageQuantities.marking ?? ""),
        status: "Inspection Done",
        rattling_scrap_qty: toNumericValue(scrapQuantities.rattling ?? ""),
        external_scrap_qty: toNumericValue(scrapQuantities.external ?? ""),
        jetting_scrap_qty: toNumericValue(scrapQuantities.jetting ?? ""),
        mpi_scrap_qty: toNumericValue(scrapQuantities.mpi ?? ""),
        drift_scrap_qty: toNumericValue(scrapQuantities.drift ?? ""),
        emi_scrap_qty: toNumericValue(scrapQuantities.emi ?? ""),
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

      toast({
        title: "Inspection updated",
        description: `${record.batch} marked as Inspection Done.`
      });

      setInitialValues({
        scrapQuantities: { ...scrapQuantities },
        class1,
        class2,
        class3,
        repair,
        scrapValue,
        startDate,
        endDate,
      });
      await refreshDataInBackground(sharePointService);
      navigate("/edit-records", {
        state: { client: record.client, wo_no: record.wo_no, batch: record.batch },
      });
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

  const isDirty = useMemo(() => {
    if (!record) {
      return false;
    }

    const quantitiesChanged = (Object.keys(initialValues.scrapQuantities) as ScrapKey[]).some(
      key => sanitizeNumberString(scrapQuantities[key] ?? "") !== sanitizeNumberString(initialValues.scrapQuantities[key] ?? "")
    );

    return (
      quantitiesChanged ||
      class1 !== initialValues.class1 ||
      class2 !== initialValues.class2 ||
      class3 !== initialValues.class3 ||
      repair !== initialValues.repair ||
      sanitizeNumberString(scrapValue) !== sanitizeNumberString(initialValues.scrapValue) ||
      startDate !== initialValues.startDate ||
      endDate !== initialValues.endDate
    );
  }, [class1, class2, class3, endDate, initialValues, record, repair, scrapQuantities, scrapValue, startDate]);

  useUnsavedChangesWarning(isDirty && !isSaving);

  const discardAndReturn = () => {
    if (!record) {
      navigate("/edit-records");
      return;
    }

    setScrapQuantities({ ...initialValues.scrapQuantities });
    setClass1(initialValues.class1);
    setClass2(initialValues.class2);
    setClass3(initialValues.class3);
    setRepair(initialValues.repair);
    setScrapValue(initialValues.scrapValue);
    setStartDate(initialValues.startDate);
    setEndDate(initialValues.endDate);
    toast({
      title: "Changes discarded",
      description: "Inspection data has been restored.",
    });

    window.setTimeout(() => {
      navigate("/edit-records", {
        state: { client: record.client, wo_no: record.wo_no, batch: record.batch },
      });
    }, 0);
  };

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
          <Card className="border border-dashed border-blue-200 bg-white">
            <CardContent className="p-6 text-center text-sm text-blue-700">
              Batch details not found. Please return to Edit Records and select a batch.
            </CardContent>
          </Card>
        ) : (
          <div className="grid gap-6 lg:grid-cols-[1.1fr,1fr]">
            <Card className="border border-blue-200 shadow-sm">
              <CardHeader className="border-b bg-white/80">
                <CardTitle className="text-lg font-semibold text-blue-900">Inspection Stages</CardTitle>
              </CardHeader>
              <CardContent className="space-y-4 p-5">
                <div className="grid gap-3 rounded-lg border border-blue-100 bg-blue-50/60 p-3 text-sm md:grid-cols-4">
                  <div>
                    <p className="text-[11px] uppercase tracking-wide text-blue-700">Client</p>
                    <p className="font-semibold text-blue-900">{record.client}</p>
                  </div>
                  <div>
                    <p className="text-[11px] uppercase tracking-wide text-blue-700">Work Order</p>
                    <p className="font-semibold text-blue-900">{record.wo_no}</p>
                  </div>
                  <div>
                    <p className="text-[11px] uppercase tracking-wide text-blue-700">Batch</p>
                    <p className="font-semibold text-blue-900">{record.batch}</p>
                  </div>
                  <div>
                    <p className="text-[11px] uppercase tracking-wide text-blue-700">Quantity</p>
                    <p className="font-semibold text-blue-900">{record.qty || "0"}</p>
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
                              value={computedStageQuantities[stage.key] ?? ""}
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

            <Card className="border border-emerald-200 shadow-sm">
              <CardHeader className="border-b bg-white/80">
                <CardTitle className="text-lg font-semibold text-emerald-900">Inspection Data</CardTitle>
              </CardHeader>
              <CardContent className="space-y-4 p-5">
                <div className="grid gap-4 md:grid-cols-2">
                  <div className="space-y-2">
                    <Label htmlFor="class1">Class 1</Label>
                    <Input id="class1" value={class1} onChange={event => setClass1(event.target.value)} placeholder="Enter Class 1" className="h-11" />
                  </div>
                  <div className="space-y-2">
                    <Label htmlFor="class2">Class 2</Label>
                    <Input id="class2" value={class2} onChange={event => setClass2(event.target.value)} placeholder="Enter Class 2" className="h-11" />
                  </div>
                  <div className="space-y-2">
                    <Label htmlFor="class3">Class 3</Label>
                    <Input id="class3" value={class3} onChange={event => setClass3(event.target.value)} placeholder="Enter Class 3" className="h-11" />
                  </div>
                  <div className="space-y-2">
                    <Label htmlFor="repair">Repair</Label>
                    <Input id="repair" value={repair} onChange={event => setRepair(event.target.value)} placeholder="Enter Repair" className="h-11" />
                  </div>
                </div>

                <div className="grid gap-4 md:grid-cols-2">
                  <DateInputField label="Start Date" value={startDate} onChange={setStartDate} className="h-11" />
                  <DateInputField label="End Date" value={endDate} onChange={setEndDate} className="h-11" />
                </div>

                <div className="grid gap-4 md:grid-cols-[2fr,1fr]">
                  <div className="space-y-2">
                    <Label htmlFor="scrap">Scrap</Label>
                    <Input
                      id="scrap"
                      value={scrapValue}
                      onChange={event => setScrapValue(sanitizeNumberString(event.target.value))}
                      placeholder="Total scrap"
                      inputMode="numeric"
                      className="h-11"
                    />
                  </div>
                  <div className="rounded-lg border border-emerald-200 bg-emerald-50/60 p-3 text-sm text-emerald-800">
                    <p className="font-semibold">Computed Scrap</p>
                    <p>{computedScrapTotal}</p>
                  </div>
                </div>

                <div className="flex flex-col items-stretch justify-end gap-2 sm:flex-row">
                  <Button
                    type="button"
                    variant="outline"
                    onClick={discardAndReturn}
                    className="h-11 min-w-[120px]"
                  >
                    Cancel
                  </Button>
                  <Button onClick={handleUpdate} disabled={isSaving} className="h-11 min-w-[160px]">
                    {isSaving ? "Updating..." : "Update"}
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
