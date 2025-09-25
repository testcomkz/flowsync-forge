import { useEffect, useMemo, useState } from "react";
import { useLocation, useNavigate } from "react-router-dom";
import { ArrowLeft, Layers } from "lucide-react";

import { Header } from "@/components/layout/Header";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { DateInputField, toDateInputValue } from "@/components/ui/date-input";
import { useToast } from "@/hooks/use-toast";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useSharePointInstantData } from "@/hooks/useInstantData";
import { computePipeTo, parseTubingRecords, sanitizeNumberString } from "@/lib/tubing-records";
import { useUnsavedChangesWarning } from "@/hooks/use-unsaved-changes";

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

  const [initialValues, setInitialValues] = useState({
    quantity: "",
    rack: "",
    arrivalDate: "",
  });
  const [quantity, setQuantity] = useState(initialValues.quantity);
  const [rack, setRack] = useState(initialValues.rack);
  const [arrivalDate, setArrivalDate] = useState(initialValues.arrivalDate);
  const [isSaving, setIsSaving] = useState(false);

  useEffect(() => {
    if (!record) {
      setInitialValues({ quantity: "", rack: "", arrivalDate: "" });
      return;
    }
    const nextInitial = {
      quantity: sanitizeNumberString(record.qty || ""),
      rack: record.rack || "",
      arrivalDate: toDateInputValue(record.arrival_date),
    };
    setInitialValues({ ...nextInitial });
    setQuantity(nextInitial.quantity);
    setRack(nextInitial.rack);
    setArrivalDate(nextInitial.arrivalDate);
  }, [record]);

  const handleBack = () => {
    if (record) {
      navigate("/edit-records", {
        state: { client: record.client, wo_no: record.wo_no, batch: record.batch },
      });
      return;
    }
    navigate("/edit-records");
  };

  const pipeFrom = record?.pipe_from ?? "";
  const computedPipeTo = useMemo(() => computePipeTo(pipeFrom, quantity || ""), [pipeFrom, quantity]);

  const isDirty =
    sanitizeNumberString(quantity) !== sanitizeNumberString(initialValues.quantity) ||
    rack !== initialValues.rack ||
    toDateInputValue(arrivalDate) !== toDateInputValue(initialValues.arrivalDate);

  useUnsavedChangesWarning(isDirty && !isSaving);

  const discardAndReturn = () => {
    if (!record) {
      navigate("/edit-records");
      return;
    }

    setQuantity(initialValues.quantity);
    setRack(initialValues.rack);
    setArrivalDate(initialValues.arrivalDate);
    toast({
      title: "Changes discarded",
      description: "Tubing registry data has been restored.",
    });

    window.setTimeout(() => {
      navigate("/edit-records", {
        state: { client: record.client, wo_no: record.wo_no, batch: record.batch },
      });
    }, 0);
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
        status: record.status
      });

      if (!success) {
        toast({
          title: "Update failed",
          description: "Unable to update tubing registry. Please try again.",
          variant: "destructive"
        });
        return;
      }

      toast({
        title: "Tubing registry updated",
        description: `${record.batch} saved successfully.`
      });

      setInitialValues({
        quantity: sanitizeNumberString(quantity),
        rack,
        arrivalDate: toDateInputValue(arrivalDate),
      });
      await refreshDataInBackground(sharePointService);
      navigate("/edit-records", {
        state: { client: record.client, wo_no: record.wo_no, batch: record.batch },
      });
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

        <Card className="border border-amber-200 shadow-sm">
          <CardHeader className="border-b bg-white/80">
            <CardTitle className="text-lg font-semibold text-amber-900">Update Tubing Registry</CardTitle>
          </CardHeader>
          <CardContent className="space-y-6 p-5">
            {missingSelection || !record ? (
              <div className="rounded-lg border border-dashed border-amber-300 bg-white p-6 text-center text-sm text-amber-700">
                Batch details not found. Please return to Edit Records and select a batch.
              </div>
            ) : (
              <>
                <div className="grid gap-3 rounded-lg border border-amber-100 bg-amber-50/60 p-3 text-sm md:grid-cols-4">
                  <div>
                    <p className="text-[11px] uppercase tracking-wide text-amber-700">Client</p>
                    <p className="font-semibold text-amber-900">{record.client}</p>
                  </div>
                  <div>
                    <p className="text-[11px] uppercase tracking-wide text-amber-700">Work Order</p>
                    <p className="font-semibold text-amber-900">{record.wo_no}</p>
                  </div>
                  <div>
                    <p className="text-[11px] uppercase tracking-wide text-amber-700">Batch</p>
                    <p className="font-semibold text-amber-900">{record.batch}</p>
                  </div>
                  <div>
                    <p className="text-[11px] uppercase tracking-wide text-amber-700">Diameter</p>
                    <p className="font-semibold text-amber-900">{record.diameter || "—"}</p>
                  </div>
                </div>

                <div className="grid gap-4 md:grid-cols-2">
                  <div className="space-y-2">
                    <Label htmlFor="quantity">Quantity</Label>
                    <Input
                      id="quantity"
                      value={quantity}
                      onChange={event => setQuantity(sanitizeNumberString(event.target.value))}
                      inputMode="numeric"
                      placeholder="Enter quantity"
                      className="h-11"
                    />
                  </div>
                  <div className="space-y-2">
                    <Label>Rack</Label>
                    <Input
                      value={rack}
                      onChange={event => setRack(event.target.value)}
                      placeholder="Enter rack"
                      className="h-11"
                    />
                  </div>
                </div>

                <div className="grid gap-4 md:grid-cols-3">
                  <div className="space-y-2">
                    <Label>Pipe From</Label>
                    <Input value={pipeFrom || "—"} readOnly className="h-11 bg-white" />
                  </div>
                  <div className="space-y-2">
                    <Label>Pipe To</Label>
                    <Input value={computedPipeTo || record?.pipe_to || "—"} readOnly className="h-11 bg-white" />
                  </div>
                  <DateInputField label="Arrival Date" value={arrivalDate} onChange={setArrivalDate} className="h-11" />
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
              </>
            )}
          </CardContent>
        </Card>
      </main>
    </div>
  );
}
