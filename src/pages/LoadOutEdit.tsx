import { useEffect, useMemo, useState } from "react";
import { useLocation, useNavigate } from "react-router-dom";
import { ArrowLeft, CheckCircle2 } from "lucide-react";

import { Header } from "@/components/layout/Header";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { DateInputField } from "@/components/ui/date-input";
import { useToast } from "@/hooks/use-toast";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useSharePointInstantData } from "@/hooks/useInstantData";
import { parseTubingRecords } from "@/lib/tubing-records";

interface LocationState {
  client?: string;
  wo_no?: string;
  batch?: string;
}

export default function LoadOutEdit() {
  const navigate = useNavigate();
  const location = useLocation();
  const { toast } = useToast();
  const { tubingData } = useSharePointInstantData();
  const { sharePointService, isConnected, refreshDataInBackground } = useSharePoint();

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

  const [loadOutDate, setLoadOutDate] = useState("");
  const [avr, setAvr] = useState("");
  const [avrDate, setAvrDate] = useState("");
  const [isSaving, setIsSaving] = useState(false);

  useEffect(() => {
    if (!record) {
      return;
    }
    setLoadOutDate(record.load_out_date || "");
    setAvr(record.act_no_oper || "");
    setAvrDate(record.act_date || "");
  }, [record]);

  const handleBack = () => {
    navigate("/edit-records");
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
      const success = await sharePointService.updateLoadOutData({
        client: record.client,
        wo_no: record.wo_no,
        batch: record.batch,
        load_out_date: loadOutDate,
        act_no_oper: avr,
        act_date: avrDate,
        status: "Completed",
        originalClient: record.originalClient,
        originalWo: record.originalWo,
        originalBatch: record.originalBatch
      });

      if (!success) {
        toast({
          title: "Update failed",
          description: "Unable to update load out data. Please try again.",
          variant: "destructive"
        });
        return;
      }

      toast({
        title: "Load Out updated",
        description: `${record.batch} marked as Completed.`
      });

      await refreshDataInBackground(sharePointService);
      navigate("/edit-records");
    } catch (error) {
      console.error("Failed to update load out data", error);
      toast({
        title: "Update failed",
        description: "Unexpected error occurred while saving load out data.",
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
          <div className="flex items-center gap-2 text-sm text-emerald-600">
            <CheckCircle2 className="h-4 w-4" />
            Load Out Edit
          </div>
        </div>

        <Card className="border-2 border-emerald-200 shadow-sm">
          <CardHeader className="border-b bg-white/80">
            <CardTitle className="text-xl font-semibold text-emerald-900">Finalize Load Out</CardTitle>
          </CardHeader>
          <CardContent className="space-y-6 p-6">
            {missingSelection || !record ? (
              <div className="rounded-lg border border-dashed border-emerald-300 bg-white p-6 text-center text-sm text-emerald-700">
                Batch details not found. Please return to Edit Records and select a batch.
              </div>
            ) : (
              <>
                <div className="grid gap-4 rounded-xl border border-emerald-100 bg-emerald-50/70 p-4 md:grid-cols-4">
                  <div>
                    <p className="text-xs uppercase tracking-wide text-emerald-700">Client</p>
                    <p className="text-base font-semibold text-emerald-900">{record.client}</p>
                  </div>
                  <div>
                    <p className="text-xs uppercase tracking-wide text-emerald-700">Work Order</p>
                    <p className="text-base font-semibold text-emerald-900">{record.wo_no}</p>
                  </div>
                  <div>
                    <p className="text-xs uppercase tracking-wide text-emerald-700">Batch</p>
                    <p className="text-base font-semibold text-emerald-900">{record.batch}</p>
                  </div>
                  <div>
                    <p className="text-xs uppercase tracking-wide text-emerald-700">Quantity</p>
                    <p className="text-base font-semibold text-emerald-900">{record.qty || "0"}</p>
                  </div>
                </div>

                <div className="grid gap-4 md:grid-cols-3">
                  <DateInputField
                    label="Load Out Date"
                    value={loadOutDate}
                    onChange={setLoadOutDate}
                  />
                  <div className="space-y-2">
                    <Label htmlFor="avr">AVR</Label>
                    <Input
                      id="avr"
                      value={avr}
                      onChange={event => setAvr(event.target.value)}
                      placeholder="Enter AVR"
                    />
                  </div>
                  <DateInputField
                    label="AVR Date"
                    value={avrDate}
                    onChange={setAvrDate}
                  />
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
