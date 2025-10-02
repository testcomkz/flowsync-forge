import { useEffect, useMemo, useRef, useState } from "react";
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
import { useSharePointInstantData } from "@/hooks/useInstantData";
import { parseTubingRecords } from "@/lib/tubing-records";
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

interface LocationState {
  client?: string;
  wo_no?: string;
  preLocks?: { hasTubingBatches: boolean; hasSuckerRodBatches: boolean; hasCouplingBatches: boolean };
  derivedDefaults?: { woType?: string; pipeType?: string };
}

export default function WorkOrderEdit() {
  const navigate = useNavigate();
  const location = useLocation();
  const { toast } = useToast();
  const { sharePointService, isConnected, refreshDataInBackground } = useSharePoint();
  const { clientRecords, tubingData, suckerRodData } = useSharePointInstantData();

  const { client: locClient, wo_no: locWO, preLocks, derivedDefaults } = (location.state as LocationState | null) ?? {};

  const [client] = useState(locClient || "");
  const [clientCode, setClientCode] = useState<string>("");
  const [woNo] = useState(locWO || "");

  const [woDate, setWoDate] = useState("");
  const [woType, setWoType] = useState(derivedDefaults?.woType ?? ""); // OCTG Inspection | Coupling Replace
  const [pipeType, setPipeType] = useState(derivedDefaults?.pipeType ?? ""); // Tubing | Sucker Rod (for OCTG)
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
  const [isOpening, setIsOpening] = useState(false);
  const [isConfirmOpen, setIsConfirmOpen] = useState(false);
  const [isCloseConfirmOpen, setIsCloseConfirmOpen] = useState(false);
  const [isOpenConfirmOpen, setIsOpenConfirmOpen] = useState(false);
  const [confirmLines, setConfirmLines] = useState<string[]>([]);
  const [originalType, setOriginalType] = useState<string>("");
  const [originalPipeType, setOriginalPipeType] = useState<string>("");
  const [originalDiameter, setOriginalDiameter] = useState<string>("");
  const [woStatus, setWoStatus] = useState<string>("");
  const isWoClosed = useMemo(() => woStatus.trim().toLowerCase().includes("closed"), [woStatus]);
  const [hasTubingBatches, setHasTubingBatches] = useState(preLocks?.hasTubingBatches ?? false);
  const [hasSuckerRodBatches, setHasSuckerRodBatches] = useState(preLocks?.hasSuckerRodBatches ?? false);
  const [hasAnyBatches, setHasAnyBatches] = useState(
    preLocks ? (preLocks.hasTubingBatches || preLocks.hasSuckerRodBatches || preLocks.hasCouplingBatches) : false
  );
  const [hasCouplingBatches, setHasCouplingBatches] = useState(preLocks?.hasCouplingBatches ?? false);

  const lastKeyRef = useRef<string | null>(null);

  const key = `${client}|${woNo}`;
  const tubingRecords = useMemo(() => parseTubingRecords(tubingData ?? []), [tubingData]);
  const suckerRecords = useMemo(() => parseTubingRecords(suckerRodData ?? []), [suckerRodData]);

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
      const rawDiameter = get("diameter");
      const rawPlannedQty = get("planned_qty") || get("planned_quantity");
      const woTypeV = get("wo_type") || get("type_of_wo");
      // Canonicalize and derive WO type
      const couplingRaw = (get("coupling_replace") || get("coupling") || get("coupling replace") || "").toString();
      const couplingYes = couplingRaw.trim().toLowerCase().startsWith("y");
      const finalWoType = (() => {
        const v = (woTypeV || "").toLowerCase();
        if (v.includes("coupling")) return "Coupling Replace";
        if (v.includes("octg")) return "OCTG Inspection";
        return couplingYes ? "Coupling Replace" : "OCTG Inspection";
      })();
      setWoType(finalWoType);
      setOriginalType(woTypeV);

      const priceTypeValue = (get("price_type") || get("pricetype") || "").trim();
      const sanitizedPricePerPipe = sanitizeDecimalInput(get("price") || get("price_for_each_pipe"));
      const sanitizedReplacement = sanitizeDecimalInput(get("replacement_price"));
      // Read WO status flexibly
      const detectStatus = () => {
        const direct = (get("wo_status") || get("status") || "").toString();
        if (direct) return direct;
        try {
          const keys = Object.keys(rec || {});
          const key = keys.find(k => k.toLowerCase().replace(/[^a-z0-9]+/g, '_').endsWith('status'));
          return key ? String((rec as any)[key] ?? '') : '';
        } catch { return ''; }
      };
      setWoStatus(detectStatus());
      // Robust fetch for Transport and Transportation Cost (headers may vary)
      const findBy = (pred: (key: string) => boolean): string => {
        try {
          const keys = Object.keys(rec || {});
          const k = keys.find(pred);
          return k ? String(rec[k] ?? "") : "";
        } catch {
          return "";
        }
      };
      const rawTransport = (
        (get("transport") || get("transport_option") ||
          findBy(k => k.includes("transport") && !k.includes("cost")))
      ).trim();
      const rawTransportCost = (
        get("transport_cost") || get("transportation_cost") ||
        findBy(k => k.includes("transport") && k.includes("cost"))
      );
      const sanitizedTransportCost = sanitizeDecimalInput(rawTransportCost);
      const canonicalTransport = (() => {
        const lower = rawTransport.toLowerCase();
        if (!lower) return "";
        if (lower.startsWith("tcc")) return "TCC";
        if (lower.startsWith("client")) return "Client";
        return rawTransport;
      })();

      if (finalWoType === "Coupling Replace") {
        setPipeType("Tubing");
        setPriceType("Fixed");
        setPricePerPipe(sanitizedPricePerPipe);
        setReplacementPrice(sanitizedReplacement);
        setTransport("");
        setTransportCost("");
      } else {
        setPipeType(typeOrPipe || "");
        setPriceType(priceTypeValue);
        setPricePerPipe(sanitizedPricePerPipe);
        setReplacementPrice(sanitizedReplacement);
        setTransport(canonicalTransport);
        // Always preload saved cost, UI will show it only when Transport is TCC
        setTransportCost(sanitizedTransportCost);
      }

      setDiameter(rawDiameter);
      setPlannedQty(rawPlannedQty);
      // Track locks inferred from existing data
      setHasCouplingBatches(couplingYes);
      if (typeOrPipe) setOriginalPipeType(typeOrPipe);
      const di = get("diameter");
      if (di) setOriginalDiameter(di);
      // coupling replace is derived from woType in UI
      // Stage prices
      const s: StagePrices = createEmptyStagePrices();
      s.rattling_price = sanitizeDecimalInput(get("rattling_price") || get("item_1_7_rattling_price"));
      s.external_price = sanitizeDecimalInput(get("external_price") || get("item_1_1_external_price"));
      s.hydro_price = sanitizeDecimalInput(get("hydro_price") || get("item_1_2_hydro_price"));
      s.mpi_price = sanitizeDecimalInput(get("mpi_price") || get("item_1_5_mpi_price"));
      s.drift_price = sanitizeDecimalInput(get("drift_price") || get("item_1_3_drift_price"));
      s.emi_price = sanitizeDecimalInput(get("emi_price") || get("item_1_4_emi_price"));
      s.marking_price = sanitizeDecimalInput(get("marking_price") || get("item_1_6_marking_price"));
      setStagePrices(s);
    })();
  }, [sharePointService, client, woNo, key]);

  useEffect(() => {
    if (!sharePointService || !client || !woNo) {
      return;
    }
    let cancelled = false;
    const detectPresence = async () => {
      try {
        const presence = await sharePointService.getBatchPresence(client, woNo);
        if (cancelled) return;
        setHasTubingBatches(presence.hasTubing);
        setHasSuckerRodBatches(presence.hasSuckerRod);
      } catch (error) {
        console.warn("Failed to fetch batch presence via SharePoint service", error);
      }
    };
    detectPresence();
    return () => {
      cancelled = true;
    };
  }, [sharePointService, client, woNo]);

  useEffect(() => {
    if (!client || !woNo) {
      setOriginalDiameter("");
      return;
    }
    const relatedTubing = tubingRecords.filter(record => record.client === client && record.wo_no === woNo);
    if (relatedTubing.length > 0) {
      const first = relatedTubing.find(r => r.diameter) ?? relatedTubing[0];
      if (first?.diameter) {
        setOriginalDiameter(first.diameter);
      }
    }
  }, [client, woNo, tubingRecords]);

  useEffect(() => {
    setHasAnyBatches(hasTubingBatches || hasSuckerRodBatches || hasCouplingBatches);
  }, [hasTubingBatches, hasSuckerRodBatches, hasCouplingBatches]);

  // Lookup clientCode separately, only when clientRecords loads
  useEffect(() => {
    if (!client || clientCode) return;
    if (!Array.isArray(clientRecords) || clientRecords.length === 0) return;
    const found = clientRecords.find(record => record.name === client);
    if (found?.clientCode) {
      setClientCode(found.clientCode);
    }
  }, [clientRecords, client, clientCode]);

  const isCouplingReplaceWO = woType === "Coupling Replace";
  const isOctgInspection = woType === "OCTG Inspection";
  const isStageBased = priceType === "Stage Based";
  const parsedPlannedQty = Number(plannedQty ?? "");
  const basePriceValid = pricePerPipe !== "" && Number(pricePerPipe) >= 0;
  const transportRequiresCost = isOctgInspection && transport === "TCC";
  const transportCostValid = !transportRequiresCost || (transportCost !== "" && Number(transportCost) >= 0);
  const stagePricesValid = Object.values(stagePrices).every(value => value !== "" && Number(value) >= 0);
  // Lock WO type if any batch exists in any relevant sheet
  const disableWoTypeSelect = hasTubingBatches || hasSuckerRodBatches || hasCouplingBatches;
  const disablePipeTypeSelect = (hasTubingBatches || hasSuckerRodBatches);
  const disableDiameterSelect = hasTubingBatches || isCouplingReplaceWO || pipeType === "Sucker Rod";
  const disablePriceTypeSelect = pipeType === "Sucker Rod";

  // Ensure Sucker Rod always uses Fixed price type (read-only) and no diameter
  useEffect(() => {
    if (pipeType === "Sucker Rod") {
      if (priceType !== "Fixed") setPriceType("Fixed");
      // Do not wipe previously entered fixed price; only ensure stage prices are cleared in UI
      setStagePrices(prev => prev); // no-op to keep state type; clearing handled on save
      if (diameter) setDiameter("");
    }
  }, [pipeType]);

  // If batches exist, auto-show the appropriate WO type as read-only in UI
  useEffect(() => {
    if (hasTubingBatches || hasSuckerRodBatches) {
      setWoType("OCTG Inspection");
      // Auto-detect pipe type from batches if not set
      if (hasTubingBatches) setPipeType("Tubing");
      if (hasSuckerRodBatches) setPipeType("Sucker Rod");
    } else if (hasCouplingBatches) {
      setWoType("Coupling Replace");
    }
  }, [hasTubingBatches, hasSuckerRodBatches, hasCouplingBatches]);

  const handleSave = async () => {
    if (!sharePointService || !isConnected) {
      toast({ title: "SharePoint not connected", variant: "destructive" });
      return;
    }
    if (!client || !woNo) {
      toast({ title: "Missing selection", description: "Open this page via selection.", variant: "destructive" });
      return;
    }

    // WO Type hard lock based on existing batches
    if ((hasTubingBatches || hasSuckerRodBatches) && woType !== "OCTG Inspection") {
      toast({ title: "Cannot change WO Type", description: "Batches exist: WO Type must be OCTG Inspection.", variant: "destructive" });
      return;
    }
    if (hasCouplingBatches && woType !== "Coupling Replace") {
      toast({ title: "Cannot change WO Type", description: "Coupling batches exist: WO Type must be Coupling Replace.", variant: "destructive" });
      return;
    }
    // Pipe Type hard lock based on existing batches
    if (hasTubingBatches && pipeType !== "Tubing") {
      toast({ title: "Cannot change Pipe Type", description: "Tubing batches exist: Pipe Type must be Tubing.", variant: "destructive" });
      return;
    }
    if (hasSuckerRodBatches && pipeType !== "Sucker Rod") {
      toast({ title: "Cannot change Pipe Type", description: "Sucker Rod batches exist: Pipe Type must be Sucker Rod.", variant: "destructive" });
      return;
    }
    if (hasTubingBatches && originalDiameter && pipeType === "Tubing" && diameter !== originalDiameter) {
      toast({ title: "Cannot change Diameter", description: "Tubing batches already exist. Diameter is locked.", variant: "destructive" });
      return;
    }

    const errors: string[] = [];
    const requireNonNegative = (value: string, label: string) => {
      const n = Number(value);
      if (value === "" || Number.isNaN(n) || n < 0) {
        errors.push(`${label} должно быть числом не меньше 0`);
      }
    };

    if (!woType) {
      errors.push("Выберите Type of WO");
    }

    if (isOctgInspection) {
      if (!pipeType) {
        errors.push("Выберите Type of Pipe");
      }
      if (!plannedQty) {
        errors.push("Planned Qty обязательно");
      } else if (!Number.isFinite(parsedPlannedQty) || parsedPlannedQty <= 0) {
        errors.push("Planned Qty должно быть больше 0");
      }
      if (!priceType) {
        errors.push("Выберите Price Type");
      }
      if (pipeType === "Sucker Rod" && priceType !== "Fixed") {
        errors.push("Для Sucker Rod Price Type всегда Fixed");
      }
      if (pipeType === "Tubing" && !diameter) {
        errors.push("Выберите Diameter");
      }

      if (priceType === "Fixed") {
        if (!basePriceValid) {
          errors.push("Price for each pipe не может быть пустым (укажите 0 или больше)");
        }
      }

      if (isStageBased) {
        stagePriceFields.forEach(field => {
          const value = stagePrices[field.key];
          if (value === "" || Number.isNaN(Number(value)) || Number(value) < 0) {
            errors.push(`${field.label} должен быть числом 0 или больше`);
          }
        });
      }
    }

    if (isCouplingReplaceWO) {
      requireNonNegative(replacementPrice, "Replacement price");
    }

    if (transportRequiresCost) {
      requireNonNegative(transportCost, "Transportation Cost");
    }

    if (errors.length) {
      toast({ title: "Нельзя сохранить", description: errors[0], variant: "destructive" });
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
      const emptyStage = createEmptyStagePrices();
      const payload: any = {
        originalClient: client,
        originalWo: woNo,
        client,
        wo_no: woNo,
        clientCode,
        // For Coupling Replace, OCTG-specific fields must be cleared
        type: isCouplingReplaceWO ? "" : pipeType,
        diameter: isCouplingReplaceWO || pipeType === "Sucker Rod" ? "" : diameter,
        coupling_replace: isCouplingReplaceWO ? "Yes" : "No",
        wo_date: woDate,
        transport: isCouplingReplaceWO ? "" : transport,
        planned_qty: isCouplingReplaceWO ? "" : plannedQty,
        wo_type: woType,
        // Price type only relevant for OCTG; clear it when not
        price_type: isCouplingReplaceWO ? "Fixed" : (isOctgInspection ? priceType : ""),
        // If switching to Stage Based -> blank price_per_pipe; if Fixed -> keep, else blank
        price_per_pipe: isCouplingReplaceWO ? pricePerPipe : (isOctgInspection && priceType === "Fixed" ? pricePerPipe : ""),
        replacement_price: isCouplingReplaceWO ? replacementPrice : "",
        transport_cost: !isCouplingReplaceWO && transport === "TCC" ? transportCost : "",
        // Always send stage prices; blank them when not Stage Based to clear Excel
        stage_prices: isOctgInspection ? (isStageBased ? stagePrices : emptyStage) : emptyStage,
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

  const doOpenWO = async () => {
    if (!sharePointService) return;
    setIsOpening(true);
    try {
      const result = await sharePointService.openWorkOrder(client, woNo);
      if (!result.success) {
        toast({ title: "Cannot open", description: result.message || "Unknown error", variant: "destructive" });
        return;
      }
      setIsOpenConfirmOpen(false); // Закрыть popup сразу
      toast({ title: "Work Order opened", description: `WO ${woNo} is now open.` });
      try {
        if (refreshDataInBackground) {
          await refreshDataInBackground(sharePointService);
        }
      } catch {}
      navigate("/edit-records");
    } catch (e) {
      console.error(e);
      toast({ title: "Open failed", description: "Unexpected error.", variant: "destructive" });
    } finally {
      setIsOpening(false);
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
            <ConfirmDialog
              open={isOpenConfirmOpen}
              title="Open Work Order?"
              description="This will mark the Work Order as Open so it can be edited and batches can be added."
              lines={[`Client: ${client}`, `WO: ${woNo}`]}
              confirmText="Open WO"
              cancelText="Cancel"
              onConfirm={doOpenWO}
              onCancel={() => setIsOpenConfirmOpen(false)}
              loading={isOpening}
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
                    <DateInputField value={woDate} onChange={setWoDate} className="h-11" disabled={isWoClosed} />
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
                    }} disabled={disableWoTypeSelect || isWoClosed}>
                      <SelectTrigger className="h-11" disabled={disableWoTypeSelect || isWoClosed}><SelectValue placeholder="Select work order type"/></SelectTrigger>
                      <SelectContent>
                        <SelectItem value="OCTG Inspection">OCTG Inspection</SelectItem>
                        <SelectItem value="Coupling Replace">Coupling Replace</SelectItem>
                      </SelectContent>
                    </Select>
                    {(disableWoTypeSelect || isWoClosed) && (
                      <p className="text-xs text-muted-foreground">WO Type locked: batches already exist for this work order.</p>
                    )}
                  </div>

                  {woType !== "Coupling Replace" && (
                    <div className="space-y-2">
                      <Label>Type Of Pipe</Label>
                      <Select
                        value={pipeType}
                        onValueChange={(v) => {
                          setPipeType(v);
                          if (v === "Sucker Rod") {
                            setDiameter("");
                            setPriceType("Fixed");
                            setStagePrices(createEmptyStagePrices());
                            setPricePerPipe("");
                          }
                        }}
                        disabled={disablePipeTypeSelect || isWoClosed}
                      >
                        <SelectTrigger className="h-11" disabled={disablePipeTypeSelect || isWoClosed}><SelectValue placeholder="Select type of pipe"/></SelectTrigger>
                        <SelectContent>
                          <SelectItem value="Tubing">Tubing</SelectItem>
                          <SelectItem value="Sucker Rod">Sucker Rod</SelectItem>
                        </SelectContent>
                      </Select>
                      {(disablePipeTypeSelect || isWoClosed) && (
                        <p className="text-xs text-muted-foreground">Pipe type locked because batches already exist for this WO.</p>
                      )}
                    </div>
                  )}

                  <div className="space-y-2">
                    <Label>Diameter</Label>
                    <Select
                      value={diameter}
                      onValueChange={setDiameter}
                      disabled={disableDiameterSelect || isWoClosed}
                    >
                      <SelectTrigger className="h-11" disabled={disableDiameterSelect || isWoClosed}><SelectValue placeholder={pipeType === 'Sucker Rod' ? 'N/A for Sucker Rod' : 'Select diameter'}/></SelectTrigger>
                      <SelectContent>
                        <SelectItem value='3 1/2"'>3 1/2"</SelectItem>
                        <SelectItem value='2 7/8"'>2 7/8"</SelectItem>
                      </SelectContent>
                    </Select>
                    {(disableDiameterSelect || isWoClosed) && pipeType === "Tubing" && (
                      <p className="text-xs text-muted-foreground">Diameter locked because tubing batches already exist.</p>
                    )}
                  </div>

                  {woType !== "Coupling Replace" && (
                    <div className="space-y-2">
                      <Label>Planned Qty</Label>
                      <Input value={plannedQty} onChange={(e) => setPlannedQty(sanitizeIntegerInput(e.target.value))} inputMode="numeric" className="h-11" disabled={isWoClosed} />
                    </div>
                  )}
                </div>
                {woType === "OCTG Inspection" && (
                  <div className="space-y-8">
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div className="space-y-2">
                        <Label>Price Type</Label>
                        <Select
                          value={priceType}
                          onValueChange={(v) => { setPriceType(v); setPricePerPipe(""); setStagePrices(createEmptyStagePrices()); }}
                          disabled={disablePriceTypeSelect || isWoClosed}
                        >
                          <SelectTrigger className="h-11" disabled={disablePriceTypeSelect || isWoClosed}><SelectValue placeholder="Select price type"/></SelectTrigger>
                          <SelectContent>
                            <SelectItem value="Fixed">Fixed</SelectItem>
                            {pipeType !== "Sucker Rod" && <SelectItem value="Stage Based">Stage Based</SelectItem>}
                          </SelectContent>
                        </Select>
                        {pipeType === "Sucker Rod" && (
                          <p className="text-xs text-muted-foreground">Price Type locked to Fixed for Sucker Rod</p>
                        )}
                        <p className="text-xs text-blue-600 font-medium">Use dot (.) for decimals</p>
                      </div>
                    </div>

                    {priceType === "Fixed" && (
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        <div className="space-y-2">
                          <Label>Price for each pipe</Label>
                          <Input value={pricePerPipe} onChange={(e) => setPricePerPipe(sanitizeDecimalInput(e.target.value))} placeholder="0.00" className="h-11" inputMode="decimal" disabled={isWoClosed} />
                        </div>
                      </div>
                    )}

                    {priceType === "Stage Based" && (
                      <div className="space-y-4">
                        <div className="border-2 border-dashed border-blue-200 rounded-lg overflow-hidden">
                          {stagePriceFields.map(field => (
                            <div key={field.key} className="grid grid-cols-1 md:grid-cols-2 gap-4 px-4 py-3 border-b border-blue-100 last:border-b-0">
                              <div className="text-sm font-semibold text-gray-700">{field.label}</div>
                              <Input value={stagePrices[field.key]} onChange={(e) => setStagePrices(prev => ({ ...prev, [field.key]: sanitizeDecimalInput(e.target.value) }))} placeholder="0.00" className="h-11" inputMode="decimal" disabled={isWoClosed} />
                            </div>
                          ))}
                        </div>
                        <p className="text-xs text-blue-600 font-medium">Prices should contain digits and dot only.</p>
                      </div>
                    )}

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div className="space-y-2">
                        <Label>Transport</Label>
                        <Select value={transport} onValueChange={(v) => { setTransport(v); if (v === 'Client') setTransportCost(''); }} disabled={isWoClosed}>
                          <SelectTrigger className="h-11" disabled={isWoClosed}><SelectValue placeholder="Select"/></SelectTrigger>
                          <SelectContent>
                            <SelectItem value="Client">Client</SelectItem>
                            <SelectItem value="TCC">TCC</SelectItem>
                          </SelectContent>
                        </Select>
                      </div>
                      {transport === "TCC" && (
                        <div className="space-y-2">
                          <Label>Transportation Cost</Label>
                          <Input value={transportCost} onChange={(e) => setTransportCost(sanitizeDecimalInput(e.target.value))} placeholder="0.00" className="h-11" inputMode="decimal" disabled={isWoClosed} />
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
                      <Input value={replacementPrice} onChange={(e) => setReplacementPrice(sanitizeDecimalInput(e.target.value))} placeholder="0.00" className="h-11" inputMode="decimal" disabled={isWoClosed} />
                    </div>
                  </div>
                )}

                <div className="flex justify-between items-center">
                  {isWoClosed ? (
                    <Button variant="outline" onClick={() => setIsOpenConfirmOpen(true)} disabled={!isConnected || isOpening} className="border-green-600 text-green-700 hover:bg-green-50">{isOpening ? 'Opening...' : 'Open Work Order'}</Button>
                  ) : (
                    <Button variant="outline" onClick={handleCloseWO} disabled={!isConnected || isClosing} className="border-red-500 text-red-600 hover:bg-red-50">{isClosing ? 'Closing...' : 'Close Work Order'}</Button>
                  )}
                  <div className="flex gap-3">
                    <Button variant="destructive" onClick={() => navigate('/edit-records')}>Cancel</Button>
                    <Button onClick={handleSave} disabled={!isConnected || isSaving || isWoClosed} className="min-w-[120px] bg-blue-600 hover:bg-blue-700 text-white">{isSaving ? 'Saving...' : 'Save'}</Button>
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
