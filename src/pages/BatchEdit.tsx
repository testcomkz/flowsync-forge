import { useEffect, useMemo, useRef, useState } from "react";
import { useNavigate } from "react-router-dom";
import { ArrowLeft, Layers } from "lucide-react";

import { Header } from "@/components/layout/Header";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue
} from "@/components/ui/select";
import { useToast } from "@/hooks/use-toast";
import { useSharePointInstantData } from "@/hooks/useInstantData";
import { computePipeTo, parseTubingRecords, TubingRecord } from "@/lib/tubing-records";
import { useSharePoint } from "@/contexts/SharePointContext";
import { safeLocalStorage } from "@/lib/safe-storage";
import { ConfirmDialog } from "@/components/ui/confirm-dialog";
import { Checkbox } from "@/components/ui/checkbox";

interface StatusOption {
  label: string;
  value: "Inspection Done" | "Arrived" | "Change Arrived";
  redirect: "/load-out-edit" | "/inspection-edit" | "/tubing-registry-edit";
  excelStatus: string;
}

const STATUS_OPTIONS: StatusOption[] = [
  { label: "Inspection Done", value: "Inspection Done", redirect: "/load-out-edit", excelStatus: "Inspection Done" },
  { label: "Arrived", value: "Arrived", redirect: "/inspection-edit", excelStatus: "Arrived" },
  { label: "Change Arrived", value: "Change Arrived", redirect: "/tubing-registry-edit", excelStatus: "" }
];

const STATUS_SEQUENCE = ["Change Arrived", "Arrived", "Inspection Done", "Completed"] as const;

const normalizeStatus = (status: string) => status.trim().toLowerCase();

const getStatusRank = (status: string) => {
  const normalized = normalizeStatus(status);
  return STATUS_SEQUENCE.findIndex(candidate => normalizeStatus(candidate) === normalized);
};

const uniqueSorted = (values: string[]) => Array.from(new Set(values.filter(Boolean))).sort((a, b) => a.localeCompare(b));

export default function BatchEdit() {
  const navigate = useNavigate();
  const { toast } = useToast();
  const { tubingData } = useSharePointInstantData();
  const { sharePointService, isConnected, refreshDataInBackground } = useSharePoint();

  const [selectedClient, setSelectedClient] = useState("");
  const [selectedWorkOrder, setSelectedWorkOrder] = useState("");
  const [selectedBatch, setSelectedBatch] = useState("");
  const [selectedStatus, setSelectedStatus] = useState<StatusOption["value"] | "">("");
  const [newQtyInput, setNewQtyInput] = useState("");
  const lastRecordIdRef = useRef<string | null>(null);
  const [qtyFocused, setQtyFocused] = useState(false);
  const [isChangingQty, setIsChangingQty] = useState(false);
  const [confirmOpen, setConfirmOpen] = useState(false);
  const [confirmLoading, setConfirmLoading] = useState(false);
  const [pendingQty, setPendingQty] = useState<number | null>(null);

  const tubingRecords = useMemo(() => parseTubingRecords(tubingData ?? []), [tubingData]);

  const availableClients = useMemo(
    () => uniqueSorted(tubingRecords.map(record => record.client)),
    [tubingRecords]
  );

  const availableWorkOrders = useMemo(
    () =>
      uniqueSorted(
        tubingRecords
          .filter(record => record.client === selectedClient)
          .map(record => record.wo_no)
      ),
    [tubingRecords, selectedClient]
  );

  const availableBatches = useMemo(
    () =>
      uniqueSorted(
        tubingRecords
          .filter(record => record.client === selectedClient && record.wo_no === selectedWorkOrder)
          .map(record => record.batch)
      ),
    [tubingRecords, selectedClient, selectedWorkOrder]
  );

  const selectedRecord: TubingRecord | null = useMemo(() => {
    if (!selectedClient || !selectedWorkOrder || !selectedBatch) {
      return null;
    }
    return (
      tubingRecords.find(
        record =>
          record.client === selectedClient &&
          record.wo_no === selectedWorkOrder &&
          record.batch === selectedBatch
      ) ?? null
    );
  }, [selectedBatch, selectedClient, selectedWorkOrder, tubingRecords]);

  // Keep Qty input in sync only when switching to a different batch (avoid resetting while typing)
  useEffect(() => {
    if (!selectedRecord) {
      lastRecordIdRef.current = null;
      setNewQtyInput("");
      setIsChangingQty(false);
      return;
    }
    // Use a STABLE id based on original keys to avoid resets on background refresh (row index can change)
    const currentId = `${selectedRecord.originalClient}||${selectedRecord.originalWo}||${selectedRecord.originalBatch}`;
    const sanitizedQty = (selectedRecord.qty ?? "").replace(/[^0-9]/g, "");
    if (lastRecordIdRef.current !== currentId) {
      // User selected a different record: initialize from record
      lastRecordIdRef.current = currentId;
      setNewQtyInput(sanitizedQty);
      setIsChangingQty(false);
      return;
    }
    // Do not override user's input for the same batch (prevents blur-reset issue)
  }, [selectedRecord]);

  const sanitizeDigits = (s: string) => s.replace(/[^0-9]/g, "");
  const toNum = (s?: string | null): number | null => {
    if (s == null) return null;
    const t = String(s).replace(/[^0-9.-]/g, "");
    if (!t || /^[.-]+$/.test(t)) return null;
    const n = Number(t);
    return Number.isFinite(n) ? n : null;
  };

  // Derive existing per-stage scrap amounts from stage quantities (monotonic non-increasing assumption)
  const deriveStageScrap = (rec: TubingRecord) => {
    const r = toNum(rec.quantities.rattling) ?? toNum(rec.qty) ?? 0;
    const ext = toNum(rec.quantities.external) ?? r;
    const hyd = toNum(rec.quantities.hydro) ?? ext;
    const mp = toNum(rec.quantities.mpi) ?? hyd;
    const dr = toNum(rec.quantities.drift) ?? mp;
    const em = toNum(rec.quantities.emi) ?? dr;
    const mark = toNum(rec.quantities.marking) ?? em;
    return {
      rattling: Math.max(0, r - ext),
      external: Math.max(0, ext - hyd),
      jetting: Math.max(0, hyd - mp),
      mpi: Math.max(0, mp - dr),
      drift: Math.max(0, dr - em),
      emi: Math.max(0, em - (mark ?? 0)),
    };
  };

  // Validate new qty against existing stage scraps
  const validateNewQtyAgainstStages = (rec: TubingRecord, newQty: number) => {
    const s = deriveStageScrap(rec);
    const cumulative = [
      s.rattling,
      s.rattling + s.external,
      s.rattling + s.external + s.jetting,
      s.rattling + s.external + s.jetting + s.mpi,
      s.rattling + s.external + s.jetting + s.mpi + s.drift,
      s.rattling + s.external + s.jetting + s.mpi + s.drift + s.emi,
    ];
    const minQty = Math.max(1, ...cumulative);
    return { ok: newQty >= minQty, minQty } as const;
  };

  // computePipeTo imported from lib/tubing-records

  const availableStatusOptions = useMemo(() => {
    if (!selectedRecord) {
      return STATUS_OPTIONS;
    }

    const currentRank = getStatusRank(selectedRecord.status ?? "");
    if (currentRank === -1) {
      return STATUS_OPTIONS;
    }

    return STATUS_OPTIONS.filter(option => {
      const optionRank = getStatusRank(option.value);
      return optionRank !== -1 && optionRank < currentRank;
    });
  }, [selectedRecord]);

  useEffect(() => {
    setSelectedWorkOrder("");
    setSelectedBatch("");
  }, [selectedClient]);

  useEffect(() => {
    setSelectedBatch("");
  }, [selectedWorkOrder]);

  const handleContinue = () => {
    if (!selectedRecord) {
      toast({ title: "Select a batch", description: "Choose Client, Work Order and Batch.", variant: "destructive" });
      return;
    }
    if (!selectedStatus) {
      toast({ title: "Select action", description: "Choose an edit action.", variant: "destructive" });
      return;
    }
    const statusConfig = STATUS_OPTIONS.find(option => option.value === selectedStatus);
    if (!statusConfig) return;

    navigate(statusConfig.redirect, {
      state: { client: selectedRecord.client, wo_no: selectedRecord.wo_no, batch: selectedRecord.batch }
    });
  };

  const handleOnlyChangeQty = () => {
    if (!selectedRecord) {
      toast({ title: "Select a batch", description: "Choose Client, Work Order and Batch.", variant: "destructive" });
      return;
    }
    const sanitized = sanitizeDigits(newQtyInput);
    const newQty = Number(sanitized);
    if (!sanitized || !Number.isFinite(newQty) || newQty <= 0) {
      toast({ title: "Validation error", description: "Enter a valid Qty (positive integer)", variant: "destructive" });
      return;
    }
    // If status is Arrived - confirm and update directly (no redirect). Otherwise validate and go to Inspection Edit for review.
    const statusNorm = (selectedRecord.status || "").toLowerCase();
    if (statusNorm.includes("arriv")) {
      if (!sharePointService || !isConnected) {
        toast({ title: "SharePoint not connected", description: "Connect to SharePoint before updating records.", variant: "destructive" });
        return;
      }
      setPendingQty(newQty);
      setConfirmOpen(true);
      return;
    }
    // Not Arrived: redirect to Inspection Edit for review before saving
    const result = validateNewQtyAgainstStages(selectedRecord, newQty);
    if (!result.ok) {
      toast({ title: "Qty conflict", description: `Minimum allowed Qty is ${result.minQty} based on existing stage scraps.`, variant: "destructive" });
      return;
    }
    navigate("/inspection-edit", { state: { client: selectedRecord.client, wo_no: selectedRecord.wo_no, batch: selectedRecord.batch, overrideQty: newQty } });
  };

  const confirmSaveArrived = async () => {
    if (!selectedRecord || pendingQty == null || !sharePointService) return;
    setConfirmLoading(true);
    try {
      const pipeToVal = computePipeTo(selectedRecord.pipe_from || "", String(pendingQty)) || selectedRecord.pipe_to || "";
      const success = await sharePointService.updateTubingRecord({
        originalClient: selectedRecord.originalClient,
        originalWo: selectedRecord.originalWo,
        originalBatch: selectedRecord.originalBatch,
        client: selectedRecord.client,
        wo_no: selectedRecord.wo_no,
        batch: selectedRecord.batch,
        diameter: selectedRecord.diameter,
        qty: String(pendingQty),
        pipe_from: selectedRecord.pipe_from,
        pipe_to: pipeToVal,
        rack: selectedRecord.rack,
        arrival_date: selectedRecord.arrival_date,
        status: selectedRecord.status,
      });
      if (!success) {
        toast({ title: "Update failed", description: "Unable to update tubing registry.", variant: "destructive" });
        return;
      }
      toast({ title: "Qty updated", description: `${selectedRecord.batch} saved with new quantity.` });
      try { safeLocalStorage.removeItem("sharepoint_last_refresh"); } catch {}
      await refreshDataInBackground(sharePointService);
      setIsChangingQty(false);
      setConfirmOpen(false);
    } catch (err) {
      console.error("Failed to update Qty from Batch Selection", err);
      toast({ title: "Update failed", description: "Unexpected error occurred while saving.", variant: "destructive" });
    } finally {
      setConfirmLoading(false);
      setPendingQty(null);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50">
      <Header />
      <main className="container mx-auto px-4 py-6">
        <div className="mb-4 flex items-center justify-between">
          <Button variant="ghost" onClick={() => navigate("/edit-records")} className="flex items-center gap-2 text-slate-600">
            <ArrowLeft className="h-4 w-4" />
            Back to Edit Records
          </Button>
          <div className="flex items-center gap-2 text-sm text-blue-600"><Layers className="h-4 w-4"/> Batch Edit</div>
        </div>

        <Card className="border-2 border-blue-200 bg-white rounded-xl shadow-md">
          <CardHeader className="border-b bg-blue-50">
            <CardTitle className="text-xl font-semibold text-blue-900">Batch Selection</CardTitle>
          </CardHeader>
          <CardContent className="grid gap-5 p-5">
            <div className="grid gap-3 md:grid-cols-3">
              <div className="space-y-2">
                <Label>Client</Label>
                <Select value={selectedClient} onValueChange={setSelectedClient}>
                  <SelectTrigger>
                    <SelectValue placeholder="Select client" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableClients.map(client => (
                      <SelectItem key={client} value={client}>{client}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-2">
                <Label>Work Order</Label>
                <Select value={selectedWorkOrder} onValueChange={setSelectedWorkOrder} disabled={!selectedClient || availableWorkOrders.length === 0}>
                  <SelectTrigger>
                    <SelectValue placeholder="Select Work Order" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableWorkOrders.map(order => (
                      <SelectItem key={order} value={order}>{order}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-2">
                <Label>Batch</Label>
                <Select value={selectedBatch} onValueChange={setSelectedBatch} disabled={!selectedWorkOrder || availableBatches.length === 0}>
                  <SelectTrigger>
                    <SelectValue placeholder="Select Batch" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableBatches.map(batch => (
                      <SelectItem key={batch} value={batch}>{batch}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </div>

            {selectedRecord ? (
              <div className="grid gap-3 rounded-xl border border-blue-100 bg-white p-3 md:grid-cols-2">
                <div className="space-y-1 text-sm">
                  <p className="text-blue-700">Current Status</p>
                  <p className="text-base font-semibold text-blue-900">{selectedRecord.status || "—"}</p>
                </div>
                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <Label>Quantity</Label>
                    <label className="flex items-center gap-2 text-sm select-none">
                      <Checkbox checked={isChangingQty} onCheckedChange={v => setIsChangingQty(Boolean(v))} />
                      <span>Change Qty</span>
                    </label>
                  </div>
                  <Input
                    value={newQtyInput}
                    onChange={e => setNewQtyInput(sanitizeDigits(e.target.value))}
                    onFocus={() => setQtyFocused(true)}
                    onBlur={() => setQtyFocused(false)}
                    inputMode="numeric"
                    className="bg-white h-9"
                    placeholder="0"
                    disabled={!isChangingQty}
                  />
                  {isChangingQty && (
                    <div className="grid grid-cols-2 gap-2">
                      <div className="space-y-1">
                        <Label className="text-xs">Pipe From</Label>
                        <Input value={selectedRecord.pipe_from || ""} readOnly className="h-8 text-sm bg-gray-100" />
                      </div>
                      <div className="space-y-1">
                        <Label className="text-xs">Pipe To</Label>
                        <Input value={computePipeTo(selectedRecord.pipe_from || "", newQtyInput)} readOnly className="h-8 text-sm bg-gray-100" />
                      </div>
                    </div>
                  )}
                </div>
                <div className="space-y-2 md:col-span-2">
                  <Label>Edit Action</Label>
                  <Select value={selectedStatus} onValueChange={v => setSelectedStatus(v as StatusOption["value"])} disabled={availableStatusOptions.length === 0 || isChangingQty}>
                    <SelectTrigger>
                      <SelectValue placeholder="Select edit action" />
                    </SelectTrigger>
                    <SelectContent>
                      {availableStatusOptions.map(option => (
                        <SelectItem key={option.value} value={option.value}>{option.label}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              </div>
            ) : (
              <div className="rounded-lg border border-dashed border-blue-300 bg-white p-4 text-center text-sm text-blue-700">
                Select Client, Work Order and Batch to continue.
              </div>
            )}

            <div className="flex justify-end gap-2">
              {isChangingQty ? (
                <Button onClick={handleOnlyChangeQty} disabled={!selectedRecord} className="min-w-[160px]">Only Change Qty</Button>
              ) : (
                <Button onClick={handleContinue} disabled={!selectedRecord || !selectedStatus} className="min-w-[160px] bg-blue-600 hover:bg-blue-700 text-white">Continue to Edit</Button>
              )}
            </div>
          </CardContent>
        </Card>
        <ConfirmDialog
          open={confirmOpen}
          onCancel={() => { if (!confirmLoading) setConfirmOpen(false); }}
          onConfirm={confirmSaveArrived}
          loading={confirmLoading}
          title="Confirm Qty Change"
          lines={selectedRecord ? [
            `Client: ${selectedRecord.client}`,
            `WO: ${selectedRecord.wo_no}`,
            `Batch: ${selectedRecord.batch}`,
            `New Qty: ${pendingQty ?? newQtyInput}`
          ] : []}
          confirmText="Change Qty"
          cancelText="Cancel"
        />
      </main>
    </div>
  );
}
