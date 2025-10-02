import { useEffect, useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import { ArrowLeft, Wrench } from "lucide-react";

import { Header } from "@/components/layout/Header";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { useToast } from "@/hooks/use-toast";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useSharePointInstantData } from "@/hooks/useInstantData";
import { parseTubingRecords } from "@/lib/tubing-records";

export default function WorkOrderEditSelect() {
  const navigate = useNavigate();
  const { toast } = useToast();
  const { sharePointService } = useSharePoint();
  const { clientRecords, clients, tubingData, suckerRodData, couplingData } = useSharePointInstantData();

  const tubingRecords = useMemo(() => parseTubingRecords(tubingData ?? []), [tubingData]);
  const suckerRecords = useMemo(() => parseTubingRecords(suckerRodData ?? []), [suckerRodData]);

  const availableClients = useMemo(() => {
    const names = clientRecords.length > 0 ? clientRecords.map(r => r.name) : clients;
    return Array.from(new Set((names || []).filter(Boolean))).sort((a, b) => a.localeCompare(b));
  }, [clientRecords, clients]);

  const [selectedClient, setSelectedClient] = useState("");
  const [availableWOs, setAvailableWOs] = useState<string[]>([]);
  const [selectedWO, setSelectedWO] = useState("");

  useEffect(() => {
    if (!sharePointService || !selectedClient) {
      setAvailableWOs([]);
      setSelectedWO("");
      return;
    }
    (async () => {
      const list = await sharePointService.getWorkOrdersByClient(selectedClient);
      setAvailableWOs(list);
      setSelectedWO("");
    })();
  }, [sharePointService, selectedClient]);

  const handleContinue = async () => {
    if (!selectedClient || !selectedWO) {
      toast({ title: "Select Client and Work Order", variant: "destructive" });
      return;
    }
    try {
      let hasTubingBatches = false;
      let hasSuckerRodBatches = false;
      let hasCouplingBatches = false;

      if (sharePointService) {
        try {
          const presence = await sharePointService.getBatchPresence(selectedClient, selectedWO);
          hasTubingBatches = presence.hasTubing;
          hasSuckerRodBatches = presence.hasSuckerRod;
          hasCouplingBatches = presence.hasCoupling;
        } catch (error) {
          console.warn("Failed to detect batch presence via service, falling back to cache", error);
        }
      }

      if (!hasTubingBatches && !hasSuckerRodBatches && !hasCouplingBatches) {
        const relatedTubing = tubingRecords.filter(r => r.client === selectedClient && r.wo_no === selectedWO);
        const relatedSucker = suckerRecords.filter(r => r.client === selectedClient && r.wo_no === selectedWO);
        hasTubingBatches = relatedTubing.length > 0;
        hasSuckerRodBatches = relatedSucker.length > 0;
        if (!hasCouplingBatches && Array.isArray(couplingData) && couplingData.length > 1) {
          const headers = couplingData[0] as unknown[];
          const normalize = (value: unknown) => (value === null || value === undefined ? "" : String(value).trim().toLowerCase());
          const clientIdx = headers.findIndex(h => normalize(h).includes("client"));
          const woIdx = headers.findIndex(h => normalize(h).includes("wo"));
          if (clientIdx !== -1 && woIdx !== -1) {
            const targetClient = selectedClient.trim().toLowerCase();
            const targetWo = selectedWO.trim().toLowerCase();
            hasCouplingBatches = couplingData.slice(1).some(row => {
              const rowArr = Array.isArray(row) ? row : [];
              return normalize(rowArr[clientIdx]) === targetClient && normalize(rowArr[woIdx]) === targetWo;
            });
          }
        }
      }

      if (hasCouplingBatches) {
        toast({
          title: "This Work Order cannot be changed",
          description: "Coupling batches already exist for this Work Order. Editing is disabled.",
          variant: "destructive",
        });
        return;
      }

      if (hasTubingBatches && hasSuckerRodBatches) {
        toast({
          title: "Mixed batches detected",
          description: "Work Order содержит и Tubing, и Sucker Rod. Свяжитесь с админом для ручной обработки.",
          variant: "destructive",
        });
        return;
      }

      // Read WO record to derive coupling flag and current defaults
      const rec = sharePointService ? await sharePointService.getWorkOrderRecord(selectedClient, selectedWO) : null;
      const get = (k: string) => (rec && rec[k] != null ? String(rec[k]) : "");
      const couplingRaw = (get("coupling_replace") || get("coupling") || get("coupling replace") || "").toLowerCase();
      const couplingFlag = couplingRaw.startsWith("y");

      // Derive initial UI defaults
      const woType = couplingFlag ? "Coupling Replace" : "OCTG Inspection";
      const pipeType = hasTubingBatches
        ? "Tubing"
        : hasSuckerRodBatches
          ? "Sucker Rod"
          : (get("type") || get("pipe_type") || get("type_of_pipe"));

      navigate("/workorder-edit", {
        state: {
          client: selectedClient,
          wo_no: selectedWO,
          preLocks: { hasTubingBatches, hasSuckerRodBatches, hasCouplingBatches: couplingFlag },
          derivedDefaults: { woType, pipeType }
        }
      });
    } catch (e) {
      console.error(e);
      navigate("/workorder-edit", { state: { client: selectedClient, wo_no: selectedWO } });
    }
  };

  return (
    <div className="min-h-screen bg-slate-50">
      <Header />
      <div className="container mx-auto px-6 py-8">
        <div className="mb-6 flex items-center justify-between">
          <Button variant="ghost" onClick={() => navigate("/edit-records")} className="flex items-center gap-2 text-slate-600">
            <ArrowLeft className="w-4 h-4" />
            <span>Back to Edit Records</span>
          </Button>
          <div className="flex items-center gap-2 text-blue-600 text-sm"><Wrench className="w-4 h-4"/> Edit Work Order</div>
        </div>

        <Card className="max-w-4xl mx-auto border-2 border-blue-200 bg-white rounded-xl shadow-md">
          <CardHeader className="bg-blue-50 border-b">
            <CardTitle className="text-2xl font-bold text-blue-900">Choose Work Order</CardTitle>
          </CardHeader>
          <CardContent className="p-6 grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-2">
              <Label>Client</Label>
              <Select value={selectedClient} onValueChange={setSelectedClient}>
                <SelectTrigger className="h-11">
                  <SelectValue placeholder="Select client" />
                </SelectTrigger>
                <SelectContent>
                  {availableClients.map(c => (
                    <SelectItem key={c} value={c}>{c}</SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <div className="space-y-2">
              <Label>Work Order</Label>
              <Select value={selectedWO} onValueChange={setSelectedWO} disabled={!selectedClient || availableWOs.length === 0}>
                <SelectTrigger className="h-11">
                  <SelectValue placeholder="Select WO" />
                </SelectTrigger>
                <SelectContent>
                  {availableWOs.map(wo => (
                    <SelectItem key={wo} value={wo}>{wo}</SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <div className="md:col-span-2 flex justify-end">
              <Button onClick={handleContinue} className="min-w-[160px] bg-blue-600 hover:bg-blue-700 text-white">Continue</Button>
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
