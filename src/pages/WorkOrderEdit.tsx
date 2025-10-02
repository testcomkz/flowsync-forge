import { useEffect, useRef, useState } from "react";
import { useLocation, useNavigate } from "react-router-dom";
import { ArrowLeft, Wrench } from "lucide-react";

import { Header } from "@/components/layout/Header";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { DateInputField } from "@/components/ui/date-input";
import { useToast } from "@/hooks/use-toast";
import { useSharePoint } from "@/contexts/SharePointContext";
import { ConfirmDialog } from "@/components/ui/confirm-dialog";

// Pricing fields config
const stagePriceFields = [
  { key: "rattling_price", label: "Item: 1.7 Rattling_Price" },
  { key: "external_price", label: "Item: 1.1 External_Price" },
  { key: "hydro_price", label: "Item: 1.2 Hydro_Price" },
  { key: "mpi_price", label: "Item: 1.5 MPI_Price" },
  { key: "drift_price", label: "Item: 1.3 Drift_Price" },
  { key: "emi_price", label: "Item: 1.4 EMI_Price" },
  { key: "marking_price", label: "Item: 1.6 Marking_Price" },
] as const;

type StagePriceKey = (typeof stagePriceFields)[number]["key"];
type StagePrices = Record<StagePriceKey, string>;

const createEmptyStagePrices = (): StagePrices => ({
  rattling_price: "",
  external_price: "",
  hydro_price: "",
  mpi_price: "",
  drift_price: "",
  emi_price: "",
  marking_price: "",
});

const sanitizeDecimalInput = (input: string): string => {
  if (!input) return "";
  const normalized = input.replace(/,/g, ".");
  const filtered = normalized.replace(/[^0-9.]/g, "");
  if (filtered.startsWith(".") && filtered !== ".") {
    const after = filtered.slice(1).replace(/\./g, "");
    return after ? `0.${after}` : "0.";
  }
  const parts = filtered.split(".");
  if (parts.length === 1) return parts[0];
  const first = parts.shift() ?? "";
  const rest = parts.join("");
  const hasTrailingDotOnly = filtered.endsWith(".") && rest.length === 0;
  if (hasTrailingDotOnly) return first + ".";
  return first + (rest ? "." + rest : "");
};
const sanitizeIntegerInput = (input: string): string => input.replace(/[^0-9]/g, "");

interface LocationState { client?: string; wo_no?: string }

export default function WorkOrderEdit() {
  const navigate = useNavigate();
  const location = useLocation();
  const { toast } = useToast();
  const { sharePointService, isConnected, refreshDataInBackground } = useSharePoint();

  const { client: locClient, wo_no: locWO } = (location.state as LocationState | null) ?? {};

  const [client] = useState(locClient || "");
  const [woNo] = useState(locWO || "");

  const [woDate, setWoDate] = useState("");
  const [woType, setWoType] = useState(""); // OCTG Inspection | Coupling Replace
  const [pipeType, setPipeType] = useState(""); // Tubing | Sucker Rod (for OCTG)
  const [diameter, setDiameter] = useState("");
  const [plannedQty, setPlannedQty] = useState("");
  const [priceType, setPriceType] = useState(""); // Fixed | Stage Based
  const [pricePerPipe, setPricePerPipe] = useState("");
  const [stagePrices, setStagePrices] = useState<StagePrices>(createEmptyStagePrices());
  const [replacementPrice, setReplacementPrice] = useState("");
  const [transport, setTransport] = useState("");
  const [transportCost, setTransportCost] = useState("");

  const [isSaving, setIsSaving] = useState(false);
  const [isClosing, setIsClosing] = useState(false);
  const [isConfirmOpen, setIsConfirmOpen] = useState(false);
  const [isCloseConfirmOpen, setIsCloseConfirmOpen] = useState(false);
  const [confirmLines, setConfirmLines] = useState<string[]>([]);

  const lastKeyRef = useRef<string | null>(null);

  const key = `${client}|${woNo}`;

  // Load existing WO details only once per target WO
  useEffect(() => {
    if (!sharePointService || !client || !woNo) return;
    if (lastKeyRef.current === key) return;
    (async () => {
      const rec = await sharePointService.getWorkOrderRecord(client, woNo);
      if (!rec) return;
      lastKeyRef.current = key;
      const get = (k: string) => (rec[k] ?? "").toString();
      setWoDate(get("wo_date") || get("date"));
      const typeOrPipe = get("type") || get("pipe_type") || get("type_of_pipe");
      setPipeType(typeOrPipe || "");
      setDiameter(get("diameter"));
      setPlannedQty(get("planned_qty") || get("planned_quantity"));
      const woTypeV = get("wo_type") || get("type_of_wo");
      setWoType(woTypeV);
      const priceTypeV = get("price_type") || get("pricetype");
      setPriceType(priceTypeV);
      setPricePerPipe(get("price") || get("price_for_each_pipe"));
      setReplacementPrice(get("replacement_price"));
      setTransport(get("transport"));
      setTransportCost(get("transport_cost"));
      // coupling replace is derived from woType in UI
      // Stage prices
      const s: StagePrices = createEmptyStagePrices();
      s.rattling_price = get("rattling_price") || get("item_1_7_rattling_price");
      s.external_price = get("external_price") || get("item_1_1_external_price");
      s.hydro_price = get("hydro_price") || get("item_1_2_hydro_price");
      s.mpi_price = get("mpi_price") || get("item_1_5_mpi_price");
      s.drift_price = get("drift_price") || get("item_1_3_drift_price");
      s.emi_price = get("emi_price") || get("item_1_4_emi_price");
      s.marking_price = get("marking_price") || get("item_1_6_marking_price");
      setStagePrices(s);
    })();
  }, [sharePointService, client, woNo, key]);

  const isCouplingReplaceWO = woType === "Coupling Replace";
  const isOctgInspection = woType === "OCTG Inspection";
  const isStageBased = priceType === "Stage Based";

  const handleSave = async () => {
    if (!sharePointService || !isConnected) {
      toast({ title: "SharePoint not connected", variant: "destructive" });
      return;
    }
    if (!client || !woNo) {
      toast({ title: "Missing selection", description: "Open this page via selection.", variant: "destructive" });
      return;
    }

    // Форматировать дату правильно (если это Excel serial number, преобразовать)
    const formatDate = (dateStr: string) => {
      if (!dateStr) return '—';
      // Если это уже в формате dd/mm/yyyy, вернуть как есть
      if (dateStr.includes('/')) return dateStr;
      // Если это Excel serial number (число), преобразовать
      const num = parseFloat(dateStr);
      if (!isNaN(num) && num > 40000) {
        const excelEpoch = new Date(1899, 11, 30);
        const date = new Date(excelEpoch.getTime() + num * 86400000);
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = date.getFullYear();
        return `${day}/${month}/${year}`;
      }
      return dateStr;
    };

    const confirm = [
      `Client: ${client}`,
      `WO: ${woNo}`,
      `WO Type: ${woType || '—'}`,
      `Pipe Type: ${pipeType || '—'}`,
      `Diameter: ${diameter || '—'}`,
      `Coupling Replace: ${isCouplingReplaceWO ? 'Yes' : 'No'}`,
      `Date: ${formatDate(woDate)}`,
      `Transport: ${transport || '—'}`,
      `Planned Qty: ${plannedQty || '—'}`,
      `Price Type: ${priceType || '—'}`,
    ];
    setConfirmLines(confirm);
    setIsConfirmOpen(true);
  };

  const doUpdate = async () => {
    if (!sharePointService) return;
    setIsSaving(true);
    try {
      const payload: any = {
        originalClient: client,
        originalWo: woNo,
        client,
        wo_no: woNo,
        type: isCouplingReplaceWO ? "Tubing" : pipeType,
        diameter: isCouplingReplaceWO || pipeType === "Sucker Rod" ? "" : diameter,
        coupling_replace: isCouplingReplaceWO ? "Yes" : "No",
        wo_date: woDate,
        transport,
        planned_qty: plannedQty,
        wo_type: woType,
        price_type: isOctgInspection ? priceType : (isCouplingReplaceWO ? "Coupling Replace" : ""),
        price_per_pipe: isOctgInspection && priceType === "Fixed" ? pricePerPipe : "",
        replacement_price: isCouplingReplaceWO ? replacementPrice : "",
        transport_cost: transport === "TCC" ? transportCost : "",
        stage_prices: isOctgInspection && isStageBased ? stagePrices : undefined,
      };

      const ok = await sharePointService.updateWorkOrder(payload);
      if (!ok) {
        toast({ title: "Update failed", description: "Unable to update work order.", variant: "destructive" });
        return;
      }
      setIsConfirmOpen(false); // Закрыть popup сразу после успешного сохранения
      toast({ title: "Work Order updated", description: `${woNo} saved.` });
      try {
        if (refreshDataInBackground) {
          await refreshDataInBackground(sharePointService);
        }
      } catch {}
      navigate("/edit-records");
    } catch (e) {
      console.error(e);
      toast({ title: "Update failed", description: "Unexpected error.", variant: "destructive" });
    } finally {
      setIsSaving(false);
    }
  };

  const handleCloseWO = () => {
    setIsCloseConfirmOpen(true);
  };

  const doCloseWO = async () => {
    if (!sharePointService) return;
    setIsClosing(true);
    try {
      const result = await sharePointService.closeWorkOrder(client, woNo);
      if (!result.success) {
        toast({ title: "Cannot close", description: result.message || "Unknown error", variant: "destructive" });
        return;
      }
      setIsCloseConfirmOpen(false); // Закрыть popup сразу
      toast({ title: "Work Order closed", description: `WO ${woNo} is now closed.` });
      try {
        if (refreshDataInBackground) {
          await refreshDataInBackground(sharePointService);
        }
      } catch {}
      navigate("/edit-records");
    } catch (e) {
      console.error(e);
      toast({ title: "Close failed", description: "Unexpected error.", variant: "destructive" });
    } finally {
      setIsClosing(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50">
      <Header />
      <div className="container mx-auto px-6 py-8">
        <div className="mb-6 flex items-center justify-between">
          <Button variant="ghost" onClick={() => navigate("/workorder-edit-select")} className="flex items-center gap-2 text-slate-600">
            <ArrowLeft className="w-4 h-4" />
            <span>Back to Select WO</span>
          </Button>
          <div className="flex items-center gap-2 text-blue-600 text-sm"><Wrench className="w-4 h-4"/> Edit Work Order</div>
        </div>

        <Card className="max-w-5xl mx-auto border-2 border-blue-200 bg-white rounded-xl shadow-md">
          <CardHeader className="bg-blue-50 border-b">
            <CardTitle className="text-2xl font-bold text-blue-900">Edit Work Order</CardTitle>
          </CardHeader>
          <CardContent className="p-6 space-y-8">
            <ConfirmDialog
              open={isConfirmOpen}
              title="Save Work Order?"
              description="Confirm updating the selected Work Order"
              lines={confirmLines}
              confirmText="Save"
              cancelText="Cancel"
              onConfirm={doUpdate}
              onCancel={() => setIsConfirmOpen(false)}
              loading={isSaving}
            />
            <ConfirmDialog
              open={isCloseConfirmOpen}
              title="Close Work Order?"
              description="This will mark the Work Order as Closed. You cannot close a WO without batches."
              lines={[`Client: ${client}`, `WO: ${woNo}`]}
              confirmText="Close WO"
              cancelText="Cancel"
              onConfirm={doCloseWO}
              onCancel={() => setIsCloseConfirmOpen(false)}
              loading={isClosing}
            />

            {!client || !woNo ? (
              <div className="rounded-lg border border-dashed bg-white p-6 text-center text-sm text-slate-600">
                Open this page from Edit Records → Edit Work Orders.
              </div>
            ) : (
              <>
                <div className="grid gap-3 rounded-xl border border-blue-100 bg-white p-3 md:grid-cols-3">
                  <div>
                    <p className="text-xs uppercase tracking-wide text-blue-700">Client</p>
                    <p className="text-base font-semibold text-blue-900">{client}</p>
                  </div>
                  <div>
                    <p className="text-xs uppercase tracking-wide text-blue-700">Work Order</p>
                    <p className="text-base font-semibold text-blue-900">{woNo}</p>
                  </div>
                  <div className="space-y-2">
                    <Label>Date</Label>
                    <DateInputField value={woDate} onChange={setWoDate} className="h-11" />
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-2">
                    <Label>Type of WO</Label>
                    <Select value={woType} onValueChange={(v) => {
                      setWoType(v);
                      if (v === "Coupling Replace") {
                        setPipeType("Tubing");
                        setPriceType("");
                        setPricePerPipe("");
                        setStagePrices(createEmptyStagePrices());
                        setPlannedQty(""); // Очистить Planned Qty для Coupling Replace
                        setDiameter(""); // Очистить Diameter для Coupling Replace
                      }
                    }}>
                      <SelectTrigger className="h-11"><SelectValue placeholder="Select work order type"/></SelectTrigger>
                      <SelectContent>
                        <SelectItem value="OCTG Inspection">OCTG Inspection</SelectItem>
                        <SelectItem value="Coupling Replace">Coupling Replace</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>

                  {woType !== "Coupling Replace" && (
                    <div className="space-y-2">
                      <Label>Type Of Pipe</Label>
                      <Select value={pipeType} onValueChange={(v) => { setPipeType(v); if (v === "Sucker Rod") setDiameter(""); }}>
                        <SelectTrigger className="h-11"><SelectValue placeholder="Select type of pipe"/></SelectTrigger>
                        <SelectContent>
                          <SelectItem value="Tubing">Tubing</SelectItem>
                          <SelectItem value="Sucker Rod">Sucker Rod</SelectItem>
                        </SelectContent>
                      </Select>
                    </div>
                  )}

                  <div className="space-y-2">
                    <Label>Diameter</Label>
                    <Select value={diameter} onValueChange={setDiameter} disabled={woType === "Coupling Replace" || pipeType === 'Sucker Rod'}>
                      <SelectTrigger className="h-11"><SelectValue placeholder={pipeType === 'Sucker Rod' ? 'N/A for Sucker Rod' : 'Select diameter'}/></SelectTrigger>
                      <SelectContent>
                        <SelectItem value='3 1/2"'>3 1/2"</SelectItem>
                        <SelectItem value='2 7/8"'>2 7/8"</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>

                  {woType !== "Coupling Replace" && (
                    <div className="space-y-2">
                      <Label>Planned Qty</Label>
                      <Input value={plannedQty} onChange={(e) => setPlannedQty(sanitizeIntegerInput(e.target.value))} inputMode="numeric" className="h-11" />
                    </div>
                  )}
                </div>

                {woType === "OCTG Inspection" && (
                  <div className="space-y-8">
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div className="space-y-2">
                        <Label>Price Type</Label>
                        <Select value={priceType} onValueChange={(v) => { setPriceType(v); setPricePerPipe(""); setStagePrices(createEmptyStagePrices()); }}>
                          <SelectTrigger className="h-11"><SelectValue placeholder="Select price type"/></SelectTrigger>
                          <SelectContent>
                            <SelectItem value="Fixed">Fixed</SelectItem>
                            {pipeType !== "Sucker Rod" && <SelectItem value="Stage Based">Stage Based</SelectItem>}
                          </SelectContent>
                        </Select>
                        <p className="text-xs text-blue-600 font-medium">Use dot (.) for decimals</p>
                      </div>
                    </div>

                    {priceType === "Fixed" && (
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        <div className="space-y-2">
                          <Label>Price for each pipe</Label>
                          <Input value={pricePerPipe} onChange={(e) => setPricePerPipe(sanitizeDecimalInput(e.target.value))} placeholder="0.00" className="h-11" inputMode="decimal" />
                        </div>
                      </div>
                    )}

                    {priceType === "Stage Based" && (
                      <div className="space-y-4">
                        <div className="border-2 border-dashed border-blue-200 rounded-lg overflow-hidden">
                          {stagePriceFields.map(field => (
                            <div key={field.key} className="grid grid-cols-1 md:grid-cols-2 gap-4 px-4 py-3 border-b border-blue-100 last:border-b-0">
                              <div className="text-sm font-semibold text-gray-700">{field.label}</div>
                              <Input value={stagePrices[field.key]} onChange={(e) => setStagePrices(prev => ({ ...prev, [field.key]: sanitizeDecimalInput(e.target.value) }))} placeholder="0.00" className="h-11" inputMode="decimal" />
                            </div>
                          ))}
                        </div>
                        <p className="text-xs text-blue-600 font-medium">Prices should contain digits and dot only.</p>
                      </div>
                    )}

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div className="space-y-2">
                        <Label>Transport</Label>
                        <Select value={transport} onValueChange={(v) => { setTransport(v); if (v === 'Client') setTransportCost(''); }}>
                          <SelectTrigger className="h-11"><SelectValue placeholder="Select"/></SelectTrigger>
                          <SelectContent>
                            <SelectItem value="Client">Client</SelectItem>
                            <SelectItem value="TCC">TCC</SelectItem>
                          </SelectContent>
                        </Select>
                      </div>
                      {transport === "TCC" && (
                        <div className="space-y-2">
                          <Label>Transportation Cost</Label>
                          <Input value={transportCost} onChange={(e) => setTransportCost(sanitizeDecimalInput(e.target.value))} placeholder="0.00" className="h-11" inputMode="decimal" />
                        </div>
                      )}
                    </div>
                  </div>
                )}

                {woType === "Coupling Replace" && (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div className="space-y-2">
                      <Label>Type Of Pipe</Label>
                      <Input value="Tubing" readOnly className="h-11 w-full rounded-md border border-gray-300 bg-gray-100 px-3 text-gray-600 shadow-sm" />
                    </div>
                    <div className="space-y-2">
                      <Label>Price for replacement</Label>
                      <Input value={replacementPrice} onChange={(e) => setReplacementPrice(sanitizeDecimalInput(e.target.value))} placeholder="0.00" className="h-11" inputMode="decimal" />
                    </div>
                  </div>
                )}

                <div className="flex justify-between items-center">
                  <Button variant="outline" onClick={handleCloseWO} disabled={!isConnected || isClosing} className="border-red-500 text-red-600 hover:bg-red-50">{isClosing ? 'Closing...' : 'Close Work Order'}</Button>
                  <div className="flex gap-3">
                    <Button variant="destructive" onClick={() => navigate('/edit-records')}>Cancel</Button>
                    <Button onClick={handleSave} disabled={!isConnected || isSaving} className="min-w-[120px] bg-blue-600 hover:bg-blue-700 text-white">{isSaving ? 'Saving...' : 'Save'}</Button>
                  </div>
                </div>
              </>
            )}
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
