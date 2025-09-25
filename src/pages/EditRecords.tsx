import { useEffect, useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import { ArrowLeft, ClipboardEdit } from "lucide-react";

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
import { useSharePoint } from "@/contexts/SharePointContext";

import { parseTubingRecords, TubingRecord } from "@/lib/tubing-records";

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

export default function EditRecords() {
  const navigate = useNavigate();
  const { toast } = useToast();
  const { tubingData } = useSharePointInstantData();
  const { sharePointService, isConnected, refreshDataInBackground } = useSharePoint();

  const [selectedClient, setSelectedClient] = useState("");
  const [selectedWorkOrder, setSelectedWorkOrder] = useState("");
  const [selectedBatch, setSelectedBatch] = useState("");
  const [selectedStatus, setSelectedStatus] = useState<StatusOption["value"] | "">("");
  const [isSaving, setIsSaving] = useState(false);

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

  useEffect(() => {
    setSelectedStatus("");
  }, [selectedRecord?.id]);

  useEffect(() => {
    if (!selectedStatus) {
      return;
    }

    const stillAllowed = availableStatusOptions.some(option => option.value === selectedStatus);
    if (!stillAllowed) {
      setSelectedStatus("");
    }
  }, [availableStatusOptions, selectedStatus]);

  const handleSave = async () => {
    if (!selectedRecord) {
      toast({
        title: "Select a batch",
        description: "Choose Client, Work Order and Batch before saving.",
        variant: "destructive"
      });
      return;
    }

    if (!selectedStatus) {
      toast({
        title: "Select status",
        description: "Choose a new status before continuing.",
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

    const statusConfig = STATUS_OPTIONS.find(option => option.value === selectedStatus);
    if (!statusConfig) {
      toast({
        title: "Unsupported status",
        description: "The selected status is not available for update.",
        variant: "destructive"
      });
      return;
    }

    const isAllowed = availableStatusOptions.some(option => option.value === statusConfig.value);
    if (!isAllowed) {
      toast({
        title: "Invalid status transition",
        description: "You can only move the batch to an earlier stage.",
        variant: "destructive"
      });
      return;
    }

    setIsSaving(true);
    try {
      const success = await sharePointService.updateTubingRecord({
        originalClient: selectedRecord.originalClient,
        originalWo: selectedRecord.originalWo,
        originalBatch: selectedRecord.originalBatch,
        client: selectedRecord.client,
        wo_no: selectedRecord.wo_no,
        batch: selectedRecord.batch,
        diameter: selectedRecord.diameter,
        qty: selectedRecord.qty,
        pipe_from: selectedRecord.pipe_from,
        pipe_to: selectedRecord.pipe_to,
        rack: selectedRecord.rack,
        arrival_date: selectedRecord.arrival_date,
        status: statusConfig.excelStatus
      });

      if (!success) {
        toast({
          title: "Update failed",
          description: "Unable to update batch status. Please try again.",
          variant: "destructive"
        });
        return;
      }

      toast({
        title: "Status updated",
        description: `${selectedRecord.batch} is now ${statusConfig.label}.`
      });

      await refreshDataInBackground(sharePointService);

      navigate(statusConfig.redirect, {
        state: {
          client: selectedRecord.client,
          wo_no: selectedRecord.wo_no,
          batch: selectedRecord.batch
        }
      });
    } catch (error) {
      console.error("Failed to update tubing status", error);
      toast({
        title: "Update failed",
        description: "Unexpected error occurred while updating the batch.",
        variant: "destructive"
      });
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50">
      <Header />
      <main className="container mx-auto px-4 py-6">
        <div className="mb-4 flex items-center justify-between">
          <Button variant="ghost" onClick={() => navigate(-1)} className="flex items-center gap-2 text-slate-600">
            <ArrowLeft className="h-4 w-4" />
            Back
          </Button>
          <div className="flex items-center gap-2 text-sm text-muted-foreground">
            <ClipboardEdit className="h-4 w-4 text-blue-500" />
            Edit Records Workflow
          </div>
        </div>

        <Card className="border-2 border-slate-200 shadow-sm">
          <CardHeader className="border-b bg-white/80">
            <CardTitle className="text-xl font-semibold text-slate-900">Batch Selection</CardTitle>
          </CardHeader>
          <CardContent className="grid gap-6 p-6">
            <div className="grid gap-4 md:grid-cols-3">
              <div className="space-y-2">
                <Label>Client</Label>
                <Select value={selectedClient} onValueChange={setSelectedClient}>
                  <SelectTrigger>
                    <SelectValue placeholder="Select client" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableClients.map(client => (
                      <SelectItem key={client} value={client}>
                        {client}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-2">
                <Label>Work Order</Label>
                <Select
                  value={selectedWorkOrder}
                  onValueChange={setSelectedWorkOrder}
                  disabled={!selectedClient || availableWorkOrders.length === 0}
                >
                  <SelectTrigger>
                    <SelectValue placeholder="Select Work Order" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableWorkOrders.map(order => (
                      <SelectItem key={order} value={order}>
                        {order}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-2">
                <Label>Batch</Label>
                <Select
                  value={selectedBatch}
                  onValueChange={setSelectedBatch}
                  disabled={!selectedWorkOrder || availableBatches.length === 0}
                >
                  <SelectTrigger>
                    <SelectValue placeholder="Select Batch" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableBatches.map(batch => (
                      <SelectItem key={batch} value={batch}>
                        {batch}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </div>

            {selectedRecord ? (
              <div className="grid gap-4 rounded-xl border border-slate-200 bg-slate-50/80 p-4 md:grid-cols-2">
                <div className="space-y-1 text-sm">
                  <p className="text-slate-500">Current Status</p>
                  <p className="text-base font-semibold text-slate-900">{selectedRecord.status || "—"}</p>
                </div>
                <div className="space-y-2">
                  <Label>Quantity</Label>
                  <Input value={selectedRecord.qty || "0"} readOnly className="bg-white" />
                </div>
                <div className="space-y-2 md:col-span-2">
                  <Label>New Status</Label>
                  <Select
                    value={selectedStatus}
                    onValueChange={value => setSelectedStatus(value as StatusOption["value"])}
                    disabled={availableStatusOptions.length === 0}
                  >
                    <SelectTrigger>
                      <SelectValue placeholder="Select new status" />
                    </SelectTrigger>
                    <SelectContent>
                      {availableStatusOptions.map(option => (
                        <SelectItem key={option.value} value={option.value}>
                          {option.label}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                  {availableStatusOptions.length === 0 ? (
                    <p className="pt-1 text-sm text-muted-foreground">This batch is already at the earliest stage.</p>
                  ) : null}
                </div>
              </div>
            ) : (
              <div className="rounded-lg border border-dashed border-slate-300 bg-white p-4 text-center text-sm text-slate-500">
                Select Client, Work Order and Batch to continue.
              </div>
            )}

            <div className="flex justify-end">
              <Button onClick={handleSave} disabled={!selectedRecord || !selectedStatus || isSaving} className="min-w-[160px]">
                {isSaving ? "Saving..." : "Save"}
              </Button>
            </div>
          </CardContent>
        </Card>
      </main>
    </div>
  );
}
