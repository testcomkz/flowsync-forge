import { useEffect, useMemo, useRef, useState } from "react";
import { useLocation, useNavigate } from "react-router-dom";
import { ArrowLeft, Layers } from "lucide-react";

import { Header } from "@/components/layout/Header";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { DateInputField } from "@/components/ui/date-input";
import { useToast } from "@/hooks/use-toast";
import { safeLocalStorage } from "@/lib/safe-storage";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useSharePointInstantData } from "@/hooks/useInstantData";
import { computePipeTo, parseTubingRecords, sanitizeNumberString } from "@/lib/tubing-records";
import { ConfirmDialog } from "@/components/ui/confirm-dialog";

interface LocationState {
  client?: string;
  wo_no?: string;
  batch?: string;
}

export default function TubingRegistryEdit() {
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

  const [quantity, setQuantity] = useState("");
  const [rack, setRack] = useState("");
  const [arrivalDate, setArrivalDate] = useState("");
  const [isSaving, setIsSaving] = useState(false);
  const [isConfirmOpen, setIsConfirmOpen] = useState(false);
  const [confirmLines, setConfirmLines] = useState<string[]>([]);
  const lastRecordIdRef = useRef<string | null>(null);
  const initialRef = useRef<{ quantity: string; rack: string; arrivalDate: string } | null>(null);

  useEffect(() => {
    if (!record) {
      lastRecordIdRef.current = null;
      return;
    }
    // Initialize form values only when the target record changes
    if (lastRecordIdRef.current === record.id) {
      return;
    }
    lastRecordIdRef.current = record.id;
    const q = record.qty || "";
    const r = record.rack || "";
    const a = record.arrival_date || "";
    setQuantity(q);
    setRack(r);
    setArrivalDate(a);
    initialRef.current = { quantity: q, rack: r, arrivalDate: a };
  }, [record]);

  const isDirty = useMemo(() => {
    const init = initialRef.current;
    if (!init) return false;
    return init.quantity !== quantity || init.rack !== rack || init.arrivalDate !== arrivalDate;
  }, [quantity, rack, arrivalDate]);

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
    if (initialRef.current) {
      setQuantity(initialRef.current.quantity);
      setRack(initialRef.current.rack);
      setArrivalDate(initialRef.current.arrivalDate);
    }
    navigate("/");
  };

  const pipeFrom = record?.pipe_from ?? "";
  const computedPipeTo = useMemo(() => computePipeTo(pipeFrom, quantity || ""), [pipeFrom, quantity]);

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

    // If nothing changed, just notify
    if (!isDirty) {
      toast({ title: "No changes", description: "There are no changes to update." });
      return;
    }

    // Валидация даты
    if (!arrivalDate || !arrivalDate.trim()) {
      toast({
        title: "Ошибка валидации",
        description: "Выберите дату прибытия (Arrival Date)",
        variant: "destructive"
      });
      return;
    }

    // Всегда проверять: Arrival Date не может быть позже Start/End Date (равенство допускается)
    const parseDate = (dateStr: string) => {
      if (!dateStr) return null;
      const parts = dateStr.split('/');
      if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1;
        const year = parseInt(parts[2], 10);
        const dt = new Date(year, month, day);
        return (dt.getFullYear() === year && dt.getMonth() === month && dt.getDate() === day) ? dt : null;
      }
      const d = new Date(dateStr);
      return Number.isNaN(d.getTime()) ? null : d;
    };

    const arrivalDateObj = parseDate(arrivalDate);
    const startDateObj = record.start_date ? parseDate(String(record.start_date)) : null;
    const endDateObj = record.end_date ? parseDate(String(record.end_date)) : null;
    const loadOutDateObj = record.load_out_date ? parseDate(String(record.load_out_date)) : null;
    const avrDateObj = record.act_date ? parseDate(String(record.act_date)) : null;

    if (arrivalDateObj && startDateObj && arrivalDateObj > startDateObj) {
      toast({
        title: "Ошибка валидации",
        description: "Arrival Date не может быть позже Start Date",
        variant: "destructive"
      });
      return;
    }

    if (arrivalDateObj && endDateObj && arrivalDateObj > endDateObj) {
      toast({
        title: "Ошибка валидации",
        description: "Arrival Date не может быть позже End Date",
        variant: "destructive"
      });
      return;
    }

    if (arrivalDateObj && loadOutDateObj && arrivalDateObj > loadOutDateObj) {
      toast({
        title: "Ошибка валидации",
        description: "Arrival Date не может быть позже Load Out Date",
        variant: "destructive"
      });
      return;
    }

    if (arrivalDateObj && avrDateObj && arrivalDateObj > avrDateObj) {
      toast({
        title: "Ошибка валидации",
        description: "Arrival Date не может быть позже AVR Date",
        variant: "destructive"
      });
      return;
    }

    setConfirmLines([
      `Client: ${record.client}`,
      `WO: ${record.wo_no}`,
      `Batch: ${record.batch}`,
      `Qty: ${quantity}`,
      `Rack: ${rack}`
    ]);
    setIsConfirmOpen(true);
  };

  const doUpdate = async () => {
    if (!record || !sharePointService) return;
    setIsSaving(true);
    try {
      const pipeToValue = computedPipeTo || record.pipe_to || "";
      const success = await sharePointService.updateTubingRecord({
        originalClient: record.originalClient,
        originalWo: record.originalWo,
        originalBatch: record.originalBatch,
        client: record.client,
        wo_no: record.wo_no,
        batch: record.batch,
        diameter: record.diameter,
        qty: quantity,
        pipe_from: pipeFrom,
        pipe_to: pipeToValue,
        rack,
        arrival_date: arrivalDate,
        status: record.status // keep status unchanged in edit flow
      });

      if (!success) {
        toast({
          title: "Update failed",
          description: "Unable to update tubing registry. Please try again.",
          variant: "destructive"
        });
        return;
      }

      setIsConfirmOpen(false); // Закрыть popup сразу
      toast({ title: "Tubing registry updated", description: `${record.batch} saved successfully.` });

      safeLocalStorage.removeItem("sharepoint_last_refresh");
      await refreshDataInBackground(sharePointService);
      
      // Логика перехода в зависимости от статуса
      const statusNorm = (record.status || "").toLowerCase();
      if (statusNorm.includes("inspection done") || statusNorm.includes("completed")) {
        // Переход в Inspection Edit для дальнейшего редактирования
        navigate("/inspection-edit", { 
          state: { 
            client: record.client, 
            wo_no: record.wo_no, 
            batch: record.batch 
          } 
        });
      } else {
        // Для Arrived - вернуться в Edit Records
        navigate("/edit-records");
      }
    } catch (error) {
      console.error("Failed to update tubing registry", error);
      toast({
        title: "Update failed",
        description: "Unexpected error occurred while saving tubing data.",
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
      <main className="container mx-auto px-4 py-6">
        <div className="mb-4 flex items-center justify-between">
          <Button variant="ghost" onClick={handleBack} className="flex items-center gap-2 text-slate-600">
            <ArrowLeft className="h-4 w-4" />
            Back to Edit Records
          </Button>
          <div className="flex items-center gap-2 text-sm text-blue-600">
            <Layers className="h-4 w-4" />
            Tubing Registry Edit
          </div>
        </div>

        <Card className="border-2 border-blue-200 rounded-xl shadow-md">
          <CardHeader className="border-b bg-blue-50">
            <CardTitle className="text-xl font-semibold text-blue-900">Update Tubing Registry</CardTitle>
          </CardHeader>
          <CardContent className="space-y-5 p-5">
            <ConfirmDialog
              open={isConfirmOpen}
              title="Update Tubing Registry?"
              description="Confirm updating the selected batch"
              lines={confirmLines}
              confirmText="Update"
              cancelText="Cancel"
              onConfirm={doUpdate}
              onCancel={() => setIsConfirmOpen(false)}
              loading={isSaving}
            />
            {missingSelection || !record ? (
              <div className="rounded-lg border border-dashed border-blue-300 bg-white p-6 text-center text-sm text-blue-700">
                Batch details not found. Please return to Edit Records and select a batch.
              </div>
            ) : (
              <>
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
                    <p className="text-xs uppercase tracking-wide text-blue-700">Diameter</p>
                    <p className="text-base font-semibold text-blue-900">{record.diameter || "—"}</p>
                  </div>
                </div>

                <div className="grid gap-3 md:grid-cols-2">
                  <div className="space-y-2">
                    <Label htmlFor="quantity">Quantity</Label>
                    <Input
                      id="quantity"
                      value={quantity}
                      onChange={event => setQuantity(sanitizeNumberString(event.target.value))}
                      inputMode="numeric"
                      placeholder="Enter quantity"
                      className="h-9"
                    />
                  </div>
                  <div className="space-y-2">
                    <Label>Rack</Label>
                    <Select value={rack} onValueChange={setRack}>
                      <SelectTrigger className="h-9">
                        <SelectValue placeholder="Select rack" />
                      </SelectTrigger>
                      <SelectContent>
                        {Array.from({ length: 7 }).map((_, index) => {
                          const rackLabel = `Rack-${index + 1}`;
                          return (
                            <SelectItem key={rackLabel} value={rackLabel}>
                              {rackLabel}
                            </SelectItem>
                          );
                        })}
                      </SelectContent>
                    </Select>
                  </div>
                </div>

                <div className="flex gap-3">
                  <div className="flex-1 space-y-2">
                    <Label>Pipe From</Label>
                    <Input value={pipeFrom || "—"} readOnly className="h-9 w-full rounded-md border border-gray-300 bg-gray-100 px-3 text-gray-600 shadow-sm" />
                  </div>
                  <div className="flex-1 space-y-2">
                    <Label>Pipe To</Label>
                    <Input value={computedPipeTo || record?.pipe_to || "—"} readOnly className="h-9 w-full rounded-md border border-gray-300 bg-gray-100 px-3 text-gray-600 shadow-sm" />
                  </div>
                  <div className="flex-1 space-y-2">
                    <Label>Arrival Date</Label>
                    <DateInputField value={arrivalDate} onChange={setArrivalDate} className="h-9" />
                  </div>
                </div>

                <div className="flex justify-end gap-3">
                  <Button variant="destructive" onClick={handleCancel} className="min-w-[120px]">Cancel</Button>
                  <Button onClick={handleUpdate} disabled={isSaving} className="min-w-[140px] bg-blue-600 hover:bg-blue-700 text-white">
                    {isSaving 
                      ? "Saving..." 
                      : (record.status || "").toLowerCase().includes("arrived") 
                        ? "Save" 
                        : "Continue to Edit"
                    }
                  </Button>
                </div>
              </>
            )}
          </CardContent>
        </Card>
      </main>
    </div>
  );
}
