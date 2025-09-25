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

      toast({ title: "Tubing registry updated", description: `${record.batch} saved successfully.` });

      safeLocalStorage.removeItem("sharepoint_last_refresh");
      await refreshDataInBackground(sharePointService);
      navigate("/edit-records");
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
          <div className="flex items-center gap-2 text-sm text-amber-600">
            <Layers className="h-4 w-4" />
            Tubing Registry Edit
          </div>
        </div>

        <Card className="border-2 border-amber-200 shadow-sm">
          <CardHeader className="border-b bg-white/80">
            <CardTitle className="text-xl font-semibold text-amber-900">Update Tubing Registry</CardTitle>
          </CardHeader>
          <CardContent className="space-y-5 p-5">
            {missingSelection || !record ? (
              <div className="rounded-lg border border-dashed border-amber-300 bg-white p-6 text-center text-sm text-amber-700">
                Batch details not found. Please return to Edit Records and select a batch.
              </div>
            ) : (
              <>
                <div className="grid gap-3 rounded-xl border border-amber-100 bg-amber-50/70 p-3 md:grid-cols-4">
                  <div>
                    <p className="text-xs uppercase tracking-wide text-amber-700">Client</p>
                    <p className="text-base font-semibold text-amber-900">{record.client}</p>
                  </div>
                  <div>
                    <p className="text-xs uppercase tracking-wide text-amber-700">Work Order</p>
                    <p className="text-base font-semibold text-amber-900">{record.wo_no}</p>
                  </div>
                  <div>
                    <p className="text-xs uppercase tracking-wide text-amber-700">Batch</p>
                    <p className="text-base font-semibold text-amber-900">{record.batch}</p>
                  </div>
                  <div>
                    <p className="text-xs uppercase tracking-wide text-amber-700">Diameter</p>
                    <p className="text-base font-semibold text-amber-900">{record.diameter || "—"}</p>
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
                    <Input value={pipeFrom || "—"} readOnly className="bg-white h-9" />
                  </div>
                  <div className="flex-1 space-y-2">
                    <Label>Pipe To</Label>
                    <Input value={computedPipeTo || record?.pipe_to || "—"} readOnly className="bg-white h-9" />
                  </div>
                  <div className="flex-1 space-y-2">
                    <Label>Arrival Date</Label>
                    <DateInputField value={arrivalDate} onChange={setArrivalDate} className="h-9" />
                  </div>
                </div>

                <div className="flex justify-end gap-3">
                  <Button variant="destructive" onClick={handleCancel} className="min-w-[120px]">Cancel</Button>
                  <Button onClick={handleUpdate} disabled={isSaving} className="min-w-[140px]">
                    {isSaving ? "Updating..." : "Update"}
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
