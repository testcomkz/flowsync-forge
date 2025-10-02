import { useEffect, useMemo, useRef, useState } from "react";
import { useLocation, useNavigate } from "react-router-dom";
import { ArrowLeft, CheckCircle2 } from "lucide-react";

import { Header } from "@/components/layout/Header";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { DateInputField } from "@/components/ui/date-input";
import { safeLocalStorage } from "@/lib/safe-storage";
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
  const lastRecordIdRef = useRef<string | null>(null);

  // Snapshot of initial values to support Cancel and unsaved-changes guard
  const initialRef = useRef<{ loadOutDate: string; avr: string; avrDate: string } | null>(null);

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
    const nextLoadOutDate = record.load_out_date || "";
    const nextAvr = record.act_no_oper || "";
    const nextAvrDate = record.act_date || "";
    setLoadOutDate(nextLoadOutDate);
    setAvr(nextAvr);
    setAvrDate(nextAvrDate);
    initialRef.current = { loadOutDate: nextLoadOutDate, avr: nextAvr, avrDate: nextAvrDate };
  }, [record]);

  const isDirty = useMemo(() => {
    const init = initialRef.current;
    if (!init) return false;
    return init.loadOutDate !== loadOutDate || init.avr !== avr || init.avrDate !== avrDate;
  }, [loadOutDate, avr, avrDate]);

  // Warn user if trying to close/refresh with unsaved changes
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
      setLoadOutDate(initialRef.current.loadOutDate);
      setAvr(initialRef.current.avr);
      setAvrDate(initialRef.current.avrDate);
    }
    navigate("/");
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

    // Date validations: Load Out Date and AVR Date must not be earlier than Arrival/Start/End
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

    const arrival = record.arrival_date ? parseDate(String(record.arrival_date)) : null;
    const start = record.start_date ? parseDate(String(record.start_date)) : null;
    const end = record.end_date ? parseDate(String(record.end_date)) : null;
    const loadOut = loadOutDate ? parseDate(loadOutDate) : null;
    const avrDt = avrDate ? parseDate(avrDate) : null;

    const err = (msg: string) => {
      toast({ title: "Ошибка валидации", description: msg, variant: "destructive" });
    };

    if (arrival && loadOut && loadOut < arrival) { err("Load Out Date не может быть раньше Arrival Date"); return; }
    if (start && loadOut && loadOut < start) { err("Load Out Date не может быть раньше Start Date"); return; }
    if (end && loadOut && loadOut < end) { err("Load Out Date не может быть раньше End Date"); return; }

    if (arrival && avrDt && avrDt < arrival) { err("AVR Date не может быть раньше Arrival Date"); return; }
    if (start && avrDt && avrDt < start) { err("AVR Date не может быть раньше Start Date"); return; }
    if (end && avrDt && avrDt < end) { err("AVR Date не может быть раньше End Date"); return; }

    const confirmMsg = [
      'Are you sure you want to update Load Out?',
      `Client: ${record.client}`,
      `WO: ${record.wo_no}`,
      `Batch: ${record.batch}`
    ].join('\n');
    if (!window.confirm(confirmMsg)) {
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

      safeLocalStorage.removeItem("sharepoint_last_refresh");
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
          <div className="flex items-center gap-2 text-sm text-blue-600">
            <CheckCircle2 className="h-4 w-4" />
            Load Out Edit
          </div>
        </div>

        <Card className="border-2 border-blue-200 rounded-xl shadow-md">
          <CardHeader className="border-b bg-blue-50">
            <CardTitle className="text-xl font-semibold text-blue-900">Finalize Load Out</CardTitle>
          </CardHeader>
          <CardContent className="space-y-5 p-5">
            {missingSelection || !record ? (
              <div className="rounded-lg border border-dashed border-blue-300 bg-white p-6 text-center text-sm text-blue-700">
                Batch details not found. Please return to Edit Records and select a batch.
              </div>
            ) : (
              <>
                <div className="grid gap-4 rounded-xl border border-blue-100 bg-white p-4 md:grid-cols-4">
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

                <div className="flex gap-3">
                  <div className="flex-1 space-y-2">
                    <Label>Load Out Date</Label>
                    <DateInputField value={loadOutDate} onChange={setLoadOutDate} className="h-9" />
                  </div>
                  <div className="flex-1 space-y-2">
                    <Label htmlFor="avr">AVR</Label>
                    <Input
                      id="avr"
                      value={avr}
                      onChange={event => setAvr(event.target.value)}
                      placeholder="Enter AVR"
                      className="h-9"
                    />
                  </div>
                  <div className="flex-1 space-y-2">
                    <Label>AVR Date</Label>
                    <DateInputField value={avrDate} onChange={setAvrDate} className="h-9" />
                  </div>
                </div>

                <div className="flex justify-end gap-3">
                  <Button variant="destructive" onClick={handleCancel} className="min-w-[120px]">Cancel</Button>
                  <Button onClick={handleUpdate} disabled={isSaving} className="min-w-[140px] bg-blue-600 hover:bg-blue-700 text-white">
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
