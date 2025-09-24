import { useEffect, useMemo, useState } from "react";
import { useLocation, useNavigate } from "react-router-dom";
import { ArrowLeft, Layers } from "lucide-react";

import { Header } from "@/components/layout/Header";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { DateInputField } from "@/components/ui/date-input";
import { useToast } from "@/hooks/use-toast";
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

  useEffect(() => {
    if (!record) {
      return;
    }
    setQuantity(record.qty || "");
    setRack(record.rack || "");
    setArrivalDate(record.arrival_date || "");
  }, [record]);

  const handleBack = () => {
    navigate("/edit-records");
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
          <CardContent className="space-y-6 p-6">
            {missingSelection || !record ? (
              <div className="rounded-lg border border-dashed border-amber-300 bg-white p-6 text-center text-sm text-amber-700">
                Batch details not found. Please return to Edit Records and select a batch.
              </div>
            ) : (
              <>
                <div className="grid gap-4 rounded-xl border border-amber-100 bg-amber-50/70 p-4 md:grid-cols-4">
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

                <div className="grid gap-4 md:grid-cols-2">
                  <div className="space-y-2">
                    <Label htmlFor="quantity">Quantity</Label>
                    <Input
                      id="quantity"
                      value={quantity}
                      onChange={event => setQuantity(sanitizeNumberString(event.target.value))}
                      inputMode="numeric"
                      placeholder="Enter quantity"
                    />
                  </div>
                  <div className="space-y-2">
                    <Label>Rack</Label>
                    <Input
                      value={rack}
                      onChange={event => setRack(event.target.value)}
                      placeholder="Enter rack"
                    />
                  </div>
                </div>

                <div className="grid gap-4 md:grid-cols-3">
                  <div className="space-y-2">
                    <Label>Pipe From</Label>
                    <Input value={pipeFrom || "—"} readOnly className="bg-white" />
                  </div>
                  <div className="space-y-2">
                    <Label>Pipe To</Label>
                    <Input value={computedPipeTo || record?.pipe_to || "—"} readOnly className="bg-white" />
                  </div>
                  <DateInputField label="Arrival Date" value={arrivalDate} onChange={setArrivalDate} />
                </div>

                <div className="flex justify-end">
                  <Button onClick={handleUpdate} disabled={isSaving} className="min-w-[160px]">
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
