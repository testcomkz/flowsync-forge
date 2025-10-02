import { useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import { ArrowLeft, Users } from "lucide-react";

import { Header } from "@/components/layout/Header";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { useToast } from "@/hooks/use-toast";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useSharePointInstantData } from "@/hooks/useInstantData";
import { ConfirmDialog } from "@/components/ui/confirm-dialog";
import { safeLocalStorage } from "@/lib/safe-storage";

export default function EditClients() {
  const navigate = useNavigate();
  const { toast } = useToast();
  const { sharePointService, isConnected, refreshDataInBackground } = useSharePoint();
  const { clientRecords } = useSharePointInstantData();

  const sorted = useMemo(() => {
    return [...clientRecords]
      .filter(r => r.name)
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [clientRecords]);

  const [selectedOriginalName, setSelectedOriginalName] = useState<string>("");
  const [selectedClientCode, setSelectedClientCode] = useState<string>("");
  const [name, setName] = useState("");
  const [payer, setPayer] = useState("");
  const [isSaving, setIsSaving] = useState(false);

  const [isConfirmOpen, setIsConfirmOpen] = useState(false);
  const [confirmLines, setConfirmLines] = useState<string[]>([]);


  const handleSelect = (n: string) => {
    setSelectedOriginalName(n);
    const rec = sorted.find(r => r.name === n);
    setName(rec?.name || "");
    setPayer(rec?.payer || "");
    setSelectedClientCode(rec?.clientCode || "");
    // inputs are initialized only on selection to avoid resetting while typing
  };

  const doUpdate = async () => {
    if (!sharePointService || !isConnected) {
      toast({ title: "SharePoint not connected", variant: "destructive" });
      return;
    }
    setIsSaving(true);
    try {
      const ok = await sharePointService.updateClientRecord({
        originalName: selectedOriginalName,
        name: name.trim(),
        payer: payer.trim(),
      });
      if (!ok) {
        toast({ title: "Update failed", description: "Unable to update client.", variant: "destructive" });
        return;
      }

      // refresh caches
      const updatedRecords = await sharePointService.getClientRecordsFromExcel();
      const updatedNames = Array.from(new Set(updatedRecords.map(r => r.name).filter(Boolean))).sort((a,b) => a.localeCompare(b));
      const timestamp = new Date().toISOString();

      safeLocalStorage.setJSON("sharepoint_cached_client_records", updatedRecords);
      safeLocalStorage.setItem("sharepoint_cached_client_records_timestamp", timestamp);
      safeLocalStorage.dispatchStorageEvent("sharepoint_cached_client_records", JSON.stringify(updatedRecords));

      safeLocalStorage.setJSON("sharepoint_cached_clients", updatedNames);
      safeLocalStorage.setItem("sharepoint_cached_clients_timestamp", timestamp);
      safeLocalStorage.dispatchStorageEvent("sharepoint_cached_clients", JSON.stringify(updatedNames));

      toast({ title: "Client updated", description: `${selectedOriginalName} → ${name}` });

      try {
        if (refreshDataInBackground) {
          safeLocalStorage.removeItem("sharepoint_last_refresh");
          await refreshDataInBackground(sharePointService);
        }
      } catch {}
    } catch (e) {
      console.error(e);
      toast({ title: "Update failed", description: "Unexpected error.", variant: "destructive" });
    } finally {
      setIsSaving(false);
      setIsConfirmOpen(false);
    }
  };

  const handleSave = () => {
    if (!selectedOriginalName) {
      toast({ title: "Select a client", description: "Choose a client to edit.", variant: "destructive" });
      return;
    }
    if (!name.trim()) {
      toast({ title: "Validation", description: "Client Name is required.", variant: "destructive" });
      return;
    }
    if (!payer.trim()) {
      toast({ title: "Validation", description: "Payer is required.", variant: "destructive" });
      return;
    }
    const duplicateName = sorted.some(r => r.name.toLowerCase() === name.trim().toLowerCase() && r.name !== selectedOriginalName);
    if (duplicateName) {
      toast({ title: "Duplicate", description: "Another client with this name already exists.", variant: "destructive" });
      return;
    }
    setConfirmLines([
      `Old Client: ${selectedOriginalName}`,
      `New Client: ${name.trim()}`,
      `Payer: ${payer.trim() || '—'}`,
    ]);
    setIsConfirmOpen(true);
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
          <div className="flex items-center gap-2 text-blue-600 text-sm"><Users className="w-4 h-4"/> Edit Clients</div>
        </div>

        <Card className="max-w-5xl mx-auto border-2 border-blue-200 rounded-xl shadow-md">
          <CardHeader className="bg-blue-50 border-b">
            <CardTitle className="text-2xl font-bold text-blue-800">Edit Clients</CardTitle>
          </CardHeader>
          <CardContent className="p-6 grid md:grid-cols-[1.1fr,1fr] gap-6">
            <ConfirmDialog
              open={isConfirmOpen}
              title="Save Changes?"
              description="Confirm client update"
              lines={confirmLines}
              confirmText="Save"
              cancelText="Cancel"
              onConfirm={doUpdate}
              onCancel={() => setIsConfirmOpen(false)}
              loading={isSaving}
            />

            <div className="rounded-lg border bg-white divide-y">
              {sorted.length === 0 ? (
                <div className="p-4 text-sm text-gray-500">No clients</div>
              ) : (
                sorted.map(item => (
                  <button
                    key={item.name}
                    onClick={() => handleSelect(item.name)}
                    className={`w-full text-left p-3 hover:bg-blue-50 ${selectedOriginalName === item.name ? 'bg-blue-100' : ''}`}
                  >
                    <div className="flex items-center gap-2">
                      <span className="text-xs font-mono font-semibold text-blue-700 bg-blue-50 px-2 py-1 rounded">{item.clientCode}</span>
                      <span className="font-semibold text-gray-800">{item.name}</span>
                    </div>
                    <div className="text-xs text-gray-500 mt-1">Payer: {item.payer || '—'}</div>
                  </button>
                ))
              )}
            </div>

            <form onSubmit={(e) => { e.preventDefault(); handleSave(); }} className="space-y-6">
              {selectedClientCode && (
                <div className="space-y-2">
                  <Label className="text-gray-500">Client Code (Неизменяемый)</Label>
                  <Input value={selectedClientCode} readOnly className="h-11 bg-gray-100 text-gray-600 font-mono font-semibold cursor-not-allowed" />
                </div>
              )}
              <div className="space-y-2">
                <Label>Client Name</Label>
                <Input value={name} onChange={(e) => setName(e.target.value)} placeholder="Enter client name" className="h-11" />
              </div>
              <div className="space-y-2">
                <Label>Payer</Label>
                <Input value={payer} onChange={(e) => setPayer(e.target.value)} placeholder="Enter payer" className="h-11" />
              </div>

              <div className="flex justify-end gap-3">
                <Button type="button" variant="destructive" onClick={() => navigate("/edit-records")}>Cancel</Button>
                <Button type="submit" disabled={!isConnected || isSaving}>{isSaving ? 'Saving...' : 'Save'}</Button>
              </div>
            </form>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
