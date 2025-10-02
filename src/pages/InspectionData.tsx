import { useEffect, useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import { Header } from "@/components/layout/Header";
import { Card, CardContent, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import { DateInputField } from "@/components/ui/date-input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { ArrowLeft, ClipboardCheck } from "lucide-react";
import { useToast } from "@/hooks/use-toast";
import { useSharePointInstantData } from "@/hooks/useInstantData";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useAuth } from "@/contexts/AuthContext";
import { safeLocalStorage } from '@/lib/safe-storage';
import { ConfirmDialog } from "@/components/ui/confirm-dialog";

type StageKey = "rattling" | "external" | "hydro" | "mpi" | "drift" | "emi" | "marking";
type ScrapKey = "rattling" | "external" | "jetting" | "mpi" | "drift" | "emi";

interface ArrivedBatchRow {
  key: string;
  client: string;
  wo_no: string;
  batch: string;
  status: string;
  arrival_date?: string;
  class_1?: string;
  class_2?: string;
  class_3?: string;
  repair?: string;
  scrap?: string;
  start_date?: string;
  end_date?: string;
  baseQty: number | null;
  rattling_qty: number | null;
  external_qty: number | null;
  hydro_qty: number | null;
  mpi_qty: number | null;
  drift_qty: number | null;
  emi_qty: number | null;
  marking_qty: number | null;
  pipe_from?: number | null;
  pipe_to?: number | null;
}

const STAGE_ORDER: StageKey[] = [
  "rattling",
  "external",
  "hydro",
  "mpi",
  "drift",
  "emi",
  "marking"
];

const stageMeta: {
  key: StageKey;
  label: string;
  scrapKey?: "rattling" | "external" | "jetting" | "mpi" | "drift" | "emi";
  scrapLabel?: string;
}[] = [
  { key: "rattling", label: "Rattling Qty", scrapKey: "rattling", scrapLabel: "Rattling Scrap Qty" },
  { key: "external", label: "External Qty", scrapKey: "external", scrapLabel: "External Scrap Qty" },
  { key: "hydro", label: "Hydro Qty", scrapKey: "jetting", scrapLabel: "Jetting Scrap Qty" },
  { key: "mpi", label: "MPI Qty", scrapKey: "mpi", scrapLabel: "MPI Scrap Qty" },
  { key: "drift", label: "Drift Qty", scrapKey: "drift", scrapLabel: "Drift Scrap Qty" },
  { key: "emi", label: "EMI Qty", scrapKey: "emi", scrapLabel: "EMI Scrap Qty" },
  { key: "marking", label: "Marking Qty" }
];

const normalizeString = (value: unknown) =>
  value === null || value === undefined ? "" : String(value).trim();

const toNumeric = (value: unknown): number | null => {
  if (value === null || value === undefined) return null;
  const sanitized = String(value).replace(/[^0-9.-]/g, "");
  if (!sanitized || /^[.-]+$/.test(sanitized)) {
    return null;
  }
  const num = Number(sanitized);
  return Number.isFinite(num) ? num : null;
};

const sanitizeDigits = (value: string) => value.replace(/[^0-9]/g, "");

const toDateInputValue = (value: unknown) => {
  if (value === null || value === undefined) return "";
  if (typeof value === "number" && Number.isFinite(value)) {
    const excelEpoch = Date.UTC(1899, 11, 30);
    const millis = excelEpoch + value * 86400000;
    return new Date(millis).toISOString().slice(0, 10);
  }

  const stringValue = String(value).trim();
  if (!stringValue) return "";
  const isoMatch = stringValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
  }

  const numericMatch = stringValue.match(/^\d+(?:\.\d+)?$/);
  if (numericMatch) {
    const numeric = Number(stringValue);
    if (Number.isFinite(numeric)) {
      const excelEpoch = Date.UTC(1899, 11, 30);
      const millis = excelEpoch + Math.floor(numeric) * 86400000;
      return new Date(millis).toISOString().slice(0, 10);
    }
  }

  const parsed = new Date(stringValue);
  if (!Number.isNaN(parsed.getTime())) {
    return parsed.toISOString().slice(0, 10);
  }

  return "";
};

const getPreviousStage = (stage: StageKey) => {
  const index = STAGE_ORDER.indexOf(stage);
  if (index <= 0) return null;
  return STAGE_ORDER[index - 1];
};

export default function InspectionData() {
  const navigate = useNavigate();
  const { toast } = useToast();
  const { user } = useAuth();
  const { sharePointService, isConnected, refreshDataInBackground } = useSharePoint();
  const { tubingData } = useSharePointInstantData();

  const [selectedClient, setSelectedClient] = useState("");
  const [selectedWorkOrder, setSelectedWorkOrder] = useState("");
  const [selectedBatch, setSelectedBatch] = useState("");
  const [selectedRow, setSelectedRow] = useState<ArrivedBatchRow | null>(null);
  const [class1, setClass1] = useState("");
  const [class2, setClass2] = useState("");
  const [class3, setClass3] = useState("");
  const [repairValue, setRepairValue] = useState("");
  const [scrapValue, setScrapValue] = useState("");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [scrapInputs, setScrapInputs] = useState<Record<ScrapKey, string>>({
    rattling: "",
    external: "",
    jetting: "",
    mpi: "",
    drift: "",
    emi: ""
  });
  const [initialQty, setInitialQty] = useState<number>(0);
  const [processedKeys, setProcessedKeys] = useState<string[]>([]);
  const [isSaving, setIsSaving] = useState(false);
  const [isConfirmOpen, setIsConfirmOpen] = useState(false);
  const [confirmLines, setConfirmLines] = useState<string[]>([]);
  const stagesCardId = "inspection-stages";
  // track initialization to avoid overwriting user's edits when SharePoint cache refreshes
  const [initializedRowKey, setInitializedRowKey] = useState<string | null>(null);
  // Track if inspection stages are filled to enable inspection data
  const [stagesCompleted, setStagesCompleted] = useState(false);

  const arrivedBatches = useMemo(() => {
    if (!Array.isArray(tubingData) || tubingData.length < 2) {
      return [] as ArrivedBatchRow[];
    }

    const headersRow = tubingData[0];
    if (!Array.isArray(headersRow)) {
      return [] as ArrivedBatchRow[];
    }

    const headers = headersRow as unknown[];
    const normalizeHeader = (header: unknown) =>
      header === null || header === undefined ? "" : String(header).trim().toLowerCase();
    // Normalize header to alphanumeric-only for robust matching: "Rattling Qty" => "rattlingqty"
    const normalizeKey = (header: unknown) =>
      (header === null || header === undefined ? "" : String(header).toLowerCase())
        .replace(/\s+/g, "")
        .replace(/[_-]+/g, "")
        .replace(/[^a-z0-9]/g, "");

    const findIndex = (predicate: (header: string) => boolean) =>
      headers.findIndex(header => predicate(normalizeHeader(header)));

    const clientIndex = findIndex(h => normalizeKey(h).includes("client"));
    const woIndex = findIndex(h => normalizeKey(h).includes("workorder") || h.includes("wo"));
    const batchIndex = findIndex(h => normalizeKey(h).includes("batch"));
    const statusIndex = findIndex(h => normalizeKey(h).includes("status"));
    // Try to find a base Qty column (not stage-specific). Be generous with matching.
    const baseQtyIndex = findIndex(h => {
      const k = normalizeKey(h);
      return k === "qty" || k === "quantity" || (k.includes("qty") && !k.includes("scrap") && !k.includes("rattling") && !k.includes("external") && !k.includes("hydro") && !k.includes("mpi") && !k.includes("drift") && !k.includes("emi") && !k.includes("marking"));
    });
    const class1Index = findIndex(h => normalizeKey(h).includes("class1"));
    const class2Index = findIndex(h => normalizeKey(h).includes("class2"));
    const class3Index = findIndex(h => normalizeKey(h).includes("class3"));
    const repairIndex = findIndex(h => normalizeKey(h).includes("repair"));
    const scrapIndex = findIndex(h => normalizeKey(h) === "scrap" || normalizeKey(h).endsWith("scrap"));
    const startDateIndex = findIndex(h => normalizeKey(h).includes("startdate"));
    const endDateIndex = findIndex(h => normalizeKey(h).includes("enddate"));
    const rattlingQtyIndex = findIndex(h => normalizeKey(h).includes("rattlingqty"));
    const externalQtyIndex = findIndex(h => normalizeKey(h).includes("externalqty"));
    const hydroQtyIndex = findIndex(h => normalizeKey(h).includes("hydroqty"));
    const mpiQtyIndex = findIndex(h => normalizeKey(h).includes("mpiqty"));
    const driftQtyIndex = findIndex(h => normalizeKey(h).includes("driftqty"));
    const emiQtyIndex = findIndex(h => normalizeKey(h).includes("emiqty"));
    const markingQtyIndex = findIndex(h => normalizeKey(h).includes("markingqty"));
    const pipeFromIndex = findIndex(h => normalizeKey(h).includes("pipefrom"));
    const pipeToIndex = findIndex(h => normalizeKey(h).includes("pipeto"));

    if (statusIndex === -1 || clientIndex === -1 || woIndex === -1 || batchIndex === -1) {
      return [] as ArrivedBatchRow[];
    }

    const rows: ArrivedBatchRow[] = [];

    tubingData.slice(1).forEach(item => {
      if (!Array.isArray(item)) {
        return;
      }
      const row = item as unknown[];
      const rawStatus = normalizeString(row[statusIndex]);
      const normalizedStatus = rawStatus.toLowerCase();
      if (!normalizedStatus.includes("arriv")) {
        return;
      }

      const client = normalizeString(row[clientIndex]);
      const wo_no = normalizeString(row[woIndex]);
      const batch = normalizeString(row[batchIndex]);
      // Compute a reliable base quantity for the batch:
      const rattlingBase = rattlingQtyIndex !== -1 ? toNumeric(row[rattlingQtyIndex]) : null;
      const baseFromQty = baseQtyIndex !== -1 ? toNumeric(row[baseQtyIndex]) : null;
      const pFrom = pipeFromIndex !== -1 ? toNumeric(row[pipeFromIndex]) : null;
      const pTo = pipeToIndex !== -1 ? toNumeric(row[pipeToIndex]) : null;
      const baseFromPipes = pFrom !== null && pTo !== null && pTo >= pFrom ? (pTo - pFrom + 1) : null;
      // As a final fallback, use the maximum among stage quantities if they are present in the row
      const stageCandidates: Array<number | null> = [
        rattlingQtyIndex !== -1 ? toNumeric(row[rattlingQtyIndex]) : null,
        externalQtyIndex !== -1 ? toNumeric(row[externalQtyIndex]) : null,
        hydroQtyIndex !== -1 ? toNumeric(row[hydroQtyIndex]) : null,
        mpiQtyIndex !== -1 ? toNumeric(row[mpiQtyIndex]) : null,
        driftQtyIndex !== -1 ? toNumeric(row[driftQtyIndex]) : null,
        emiQtyIndex !== -1 ? toNumeric(row[emiQtyIndex]) : null,
        markingQtyIndex !== -1 ? toNumeric(row[markingQtyIndex]) : null
      ];
      const stageMax = stageCandidates.filter(v => v !== null).length
        ? Math.max(...(stageCandidates.filter((v): v is number => v !== null)))
        : null;
      const computedBase = rattlingBase ?? baseFromQty ?? baseFromPipes ?? stageMax ?? null;

      rows.push({
        key: `${client}||${wo_no}||${batch}`,
        client,
        wo_no,
        batch,
        status: rawStatus || "Arrived",
        class_1: normalizeString(class1Index === -1 ? "" : row[class1Index]),
        class_2: normalizeString(class2Index === -1 ? "" : row[class2Index]),
        class_3: normalizeString(class3Index === -1 ? "" : row[class3Index]),
        repair: normalizeString(repairIndex === -1 ? "" : row[repairIndex]),
        scrap: normalizeString(scrapIndex === -1 ? "" : row[scrapIndex]),
        start_date: normalizeString(startDateIndex === -1 ? "" : row[startDateIndex]),
        end_date: normalizeString(endDateIndex === -1 ? "" : row[endDateIndex]),
        baseQty: computedBase,
        rattling_qty: rattlingBase,
        external_qty: toNumeric(externalQtyIndex === -1 ? null : row[externalQtyIndex]),
        hydro_qty: toNumeric(hydroQtyIndex === -1 ? null : row[hydroQtyIndex]),
        mpi_qty: toNumeric(mpiQtyIndex === -1 ? null : row[mpiQtyIndex]),
        drift_qty: toNumeric(driftQtyIndex === -1 ? null : row[driftQtyIndex]),
        emi_qty: toNumeric(emiQtyIndex === -1 ? null : row[emiQtyIndex]),
        marking_qty: toNumeric(markingQtyIndex === -1 ? null : row[markingQtyIndex]),
        pipe_from: pFrom,
        pipe_to: pTo,
      });
    });

    return rows;
  }, [tubingData]);

  const availableRows = useMemo(
    () => arrivedBatches.filter(row => !processedKeys.includes(row.key)),
    [arrivedBatches, processedKeys]
  );

  const availableClients = useMemo(() => {
    const unique = new Set<string>();
    availableRows.forEach(row => {
      if (row.client) unique.add(row.client);
    });
    return Array.from(unique);
  }, [availableRows]);

  const availableWorkOrders = useMemo(() => {
    if (!selectedClient) return [] as string[];
    const unique = new Set<string>();
    availableRows
      .filter(row => row.client === selectedClient)
      .forEach(row => {
        if (row.wo_no) unique.add(row.wo_no);
      });
    return Array.from(unique);
  }, [availableRows, selectedClient]);

  const availableBatches = useMemo(() => {
    if (!selectedClient || !selectedWorkOrder) return [] as ArrivedBatchRow[];
    return availableRows.filter(
      row => row.client === selectedClient && row.wo_no === selectedWorkOrder
    );
  }, [availableRows, selectedClient, selectedWorkOrder]);

  useEffect(() => {
    if (selectedClient && !availableClients.includes(selectedClient)) {
      setSelectedClient("");
      setSelectedWorkOrder("");
      setSelectedBatch("");
      setSelectedRow(null);
    }
  }, [availableClients, selectedClient]);

  useEffect(() => {
    if (selectedWorkOrder && !availableWorkOrders.includes(selectedWorkOrder)) {
      setSelectedWorkOrder("");
      setSelectedBatch("");
      setSelectedRow(null);
    }
  }, [availableWorkOrders, selectedWorkOrder]);

  useEffect(() => {
    if (!selectedBatch) {
      setSelectedRow(null);
      return;
    }

    const match = availableBatches.find(row => row.batch === selectedBatch);
    setSelectedRow(match ?? null);
  }, [availableBatches, selectedBatch]);

  useEffect(() => {
    if (!selectedRow) {
      return;
    }

    // ALWAYS treat general Qty (baseQty) as the source of truth.
    // Rattling Qty must equal Qty, so we prefer baseQty over rattling_qty if both exist.
    const base = selectedRow.baseQty ?? selectedRow.rattling_qty ?? null;
    const hasBase = base != null;
    setInitialQty(hasBase ? base! : 0);
    // Initialize top fields only once per selected row to keep inputs editable
    if (initializedRowKey !== selectedRow.key) {
      setClass1(selectedRow.class_1 || "");
      setClass2(selectedRow.class_2 || "");
      setClass3(selectedRow.class_3 || "");
      setRepairValue(selectedRow.repair || "");
      setScrapValue(selectedRow.scrap || "");
      setStartDate(toDateInputValue(selectedRow.start_date));
      setEndDate(toDateInputValue(selectedRow.end_date));
    }

    const diff = (a: number | null, b: number | null): string => {
      if (a === null || a === undefined || b === null || b === undefined) return "";
      const d = a - b;
      return d >= 0 ? String(d) : "";
    };

    if (initializedRowKey !== selectedRow.key) {
      setScrapInputs({
        // Since Rattling Qty must equal Qty, we use the resolved base as the left side.
        rattling: diff(hasBase ? base : null, selectedRow.external_qty),
        external: diff(selectedRow.external_qty, selectedRow.hydro_qty),
        jetting: diff(selectedRow.hydro_qty, selectedRow.mpi_qty),
        mpi: diff(selectedRow.mpi_qty, selectedRow.drift_qty),
        drift: diff(selectedRow.drift_qty, selectedRow.emi_qty),
        emi: diff(selectedRow.emi_qty, selectedRow.marking_qty)
      });
      setInitializedRowKey(selectedRow.key);
    }
  }, [selectedRow, initializedRowKey]);

  useEffect(() => {
    if (selectedBatch) {
      return;
    }

    setClass1("");
    setClass2("");
    setClass3("");
    setRepairValue("");
    setScrapValue("");
    setStartDate("");
    setEndDate("");
    setScrapInputs({ rattling: "", external: "", jetting: "", mpi: "", drift: "", emi: "" });
    setInitialQty(0);
    setInitializedRowKey(null);
  }, [selectedBatch]);

  // Removed Qty change UI and handlers from Inspection Data per requirements

  const toNum = (s: string) => (s === "" ? 0 : Number(s));
  const computedQuantities = useMemo(() => {
    const r = Number.isFinite(initialQty) ? initialQty : 0;
    const ext = Math.max(0, r - toNum(scrapInputs.rattling));
    const hyd = Math.max(0, ext - toNum(scrapInputs.external));
    const mp = Math.max(0, hyd - toNum(scrapInputs.jetting));
    const dr = Math.max(0, mp - toNum(scrapInputs.mpi));
    const em = Math.max(0, dr - toNum(scrapInputs.drift));
    const mark = Math.max(0, em - toNum(scrapInputs.emi));
    return { rattling: r, external: ext, hydro: hyd, mpi: mp, drift: dr, emi: em, marking: mark };
  }, [initialQty, scrapInputs]);

  const totalScrap = useMemo(() => {
    return (Object.values(scrapInputs) as string[]).reduce((sum, v) => {
      const n = v === "" ? 0 : Number(v);
      return Number.isFinite(n) ? sum + n : sum;
    }, 0);
  }, [scrapInputs]);

  // Check if stages are completed (all scrap fields filled; '0' is allowed)
  useEffect(() => {
    const requiredKeys: ScrapKey[] = ["rattling", "external", "jetting", "mpi", "drift", "emi"];
    const allFilled = requiredKeys.every(k => scrapInputs[k] !== "");
    setStagesCompleted(!!selectedRow && allFilled);
  }, [scrapInputs, selectedRow]);
  const handleScrapChange = (key: ScrapKey, value: string) => {
    const sanitized = sanitizeDigits(value);
    const prevQtyMap: Record<ScrapKey, number> = {
      rattling: computedQuantities.rattling,
      external: computedQuantities.external,
      jetting: computedQuantities.hydro,
      mpi: computedQuantities.mpi,
      drift: computedQuantities.drift,
      emi: computedQuantities.emi
    };
    const prevQty = prevQtyMap[key] ?? 0;
    if (sanitized !== "") {
      const num = Number(sanitized);
      if (Number.isFinite(num) && num > prevQty) {
        toast({ title: "Ошибка", description: "Scrap не может превышать количество на текущем этапе", variant: "destructive" });
        return;
      }
    }
    setScrapInputs(prev => ({ ...prev, [key]: sanitized }));
  };

  const handleSave = async () => {
    if (!user) { toast({ title: "Ошибка", description: "Пожалуйста, войдите в систему", variant: "destructive" }); return; }
    if (!sharePointService || !isConnected) { toast({ title: "Ошибка", description: "SharePoint не подключен", variant: "destructive" }); return; }
    if (!selectedRow) { toast({ title: "Ошибка", description: "Выберите партию для сохранения", variant: "destructive" }); return; }

    const stageNumbers: Record<StageKey, number> = {
      rattling: computedQuantities.rattling,
      external: computedQuantities.external,
      hydro: computedQuantities.hydro,
      mpi: computedQuantities.mpi,
      drift: computedQuantities.drift,
      emi: computedQuantities.emi,
      marking: computedQuantities.marking
    };

    for (const stage of STAGE_ORDER) {
      if (!Number.isFinite(stageNumbers[stage]) || stageNumbers[stage] < 0) {
        toast({ title: "Ошибка", description: "Количества этапов вычислены некорректно", variant: "destructive" });
        return;
      }
      const prevStage = getPreviousStage(stage);
      if (prevStage && stageNumbers[prevStage] < stageNumbers[stage]) {
        toast({ title: "Ошибка", description: "Количество на следующем этапе не может превышать предыдущее", variant: "destructive" });
        return;
      }
    }

    if (stageNumbers.rattling !== initialQty) {
      toast({ title: "Ошибка", description: "Rattling Qty должно совпадать с количеством труб партии", variant: "destructive" });
      return;
    }

    // Validate totals: Class1 + Class2 + Class3 + Repair + Total Scrap must equal batch Qty
    const c1 = Number(sanitizeDigits(class1)) || 0;
    const c2 = Number(sanitizeDigits(class2)) || 0;
    const c3 = Number(sanitizeDigits(class3)) || 0;
    const rep = Number(sanitizeDigits(repairValue)) || 0;
    if (c1 + c2 + c3 + rep + totalScrap !== initialQty) {
      toast({ title: "Ошибка", description: "Сумма Class1 + Class2 + Class3 + Repair + Scrap должна равняться Qty батча", variant: "destructive" });
      return;
    }

    if (!startDate) { toast({ title: "Ошибка", description: "Выберите Start Date", variant: "destructive" }); return; }
    if (!endDate) { toast({ title: "Ошибка", description: "Выберите End Date", variant: "destructive" }); return; }
    
    // Валидация дат
    const parseDate = (dateStr: string | number | undefined | null) => {
      if (dateStr === null || dateStr === undefined || dateStr === '') return null;
      
      // Если это число (Excel serial date)
      if (typeof dateStr === 'number' && Number.isFinite(dateStr)) {
        const excelEpoch = Date.UTC(1899, 11, 30);
        const millis = excelEpoch + dateStr * 86400000;
        return new Date(millis);
      }
      
      const str = String(dateStr).trim();
      if (!str) return null;
      
      // Формат dd/mm/yyyy
      const ddmmyyyy = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (ddmmyyyy) {
        const day = parseInt(ddmmyyyy[1], 10);
        const month = parseInt(ddmmyyyy[2], 10) - 1; // месяцы с 0
        const year = parseInt(ddmmyyyy[3], 10);
        const dt = new Date(year, month, day);
        // Проверка валидности даты
        if (dt.getFullYear() === year && dt.getMonth() === month && dt.getDate() === day) {
          return dt;
        }
        return null;
      }
      
      // ISO формат yyyy-mm-dd
      const iso = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (iso) {
        const year = parseInt(iso[1], 10);
        const month = parseInt(iso[2], 10) - 1;
        const day = parseInt(iso[3], 10);
        const dt = new Date(year, month, day);
        if (dt.getFullYear() === year && dt.getMonth() === month && dt.getDate() === day) {
          return dt;
        }
        return null;
      }
      
      // Fallback для других форматов
      const d = new Date(str);
      return Number.isNaN(d.getTime()) ? null : d;
    };

    const arrivalDateObj = parseDate(selectedRow.arrival_date);
    const startDateObj = parseDate(startDate);
    const endDateObj = parseDate(endDate);

    // Debug logging
    console.log('🔍 Date Validation:', {
      arrivalRaw: selectedRow.arrival_date,
      arrivalParsed: arrivalDateObj,
      startRaw: startDate,
      startParsed: startDateObj,
      endRaw: endDate,
      endParsed: endDateObj
    });

    if (startDateObj && endDateObj && startDateObj > endDateObj) {
      toast({ title: "Ошибка", description: "End Date не может быть раньше Start Date", variant: "destructive" });
      return;
    }

    if (arrivalDateObj && startDateObj && startDateObj < arrivalDateObj) {
      console.log('❌ Start Date < Arrival Date:', startDateObj, '<', arrivalDateObj);
      toast({ title: "Ошибка", description: "Start Date не может быть раньше Arrival Date", variant: "destructive" });
      return;
    }

    if (arrivalDateObj && endDateObj && endDateObj < arrivalDateObj) {
      console.log('❌ End Date < Arrival Date:', endDateObj, '<', arrivalDateObj);
      toast({ title: "Ошибка", description: "End Date не может быть раньше Arrival Date", variant: "destructive" });
      return;
    }

    // Показать confirmation dialog
    setConfirmLines([
      `Client: ${selectedRow.client}`,
      `WO: ${selectedRow.wo_no}`,
      `Batch: ${selectedRow.batch}`,
      `Class 1: ${c1}`,
      `Class 2: ${c2}`,
      `Class 3: ${c3}`,
      `Repair: ${rep}`,
      `Total Scrap: ${totalScrap}`,
      `Start Date: ${startDate}`,
      `End Date: ${endDate}`
    ]);
    setIsConfirmOpen(true);
  };

  const doSave = async () => {
    if (!selectedRow || !sharePointService) return;

    const c1 = Number(sanitizeDigits(class1)) || 0;
    const c2 = Number(sanitizeDigits(class2)) || 0;
    const c3 = Number(sanitizeDigits(class3)) || 0;
    const rep = Number(sanitizeDigits(repairValue)) || 0;

    const stageNumbers: Record<StageKey, number> = {
      rattling: computedQuantities.rattling,
      external: computedQuantities.external,
      hydro: computedQuantities.hydro,
      mpi: computedQuantities.mpi,
      drift: computedQuantities.drift,
      emi: computedQuantities.emi,
      marking: computedQuantities.marking
    };

    setIsSaving(true);
    try {
      const success = await sharePointService.updateTubingInspectionData({
        client: selectedRow.client,
        wo_no: selectedRow.wo_no,
        batch: selectedRow.batch,
        class_1: String(c1),
        class_2: String(c2),
        class_3: String(c3),
        repair: String(rep),
        scrap: totalScrap,
        start_date: startDate,
        end_date: endDate,
        rattling_qty: stageNumbers.rattling,
        external_qty: stageNumbers.external,
        hydro_qty: stageNumbers.hydro,
        mpi_qty: stageNumbers.mpi,
        drift_qty: stageNumbers.drift,
        emi_qty: stageNumbers.emi,
        marking_qty: stageNumbers.marking,
        // НЕ отправляем scrap_qty колонки - там формулы в Excel!
        status: "Inspection Done"
      });

      if (success) {
        setIsConfirmOpen(false); // Закрыть popup
        toast({ title: "Успешно", description: "Инспекция сохранена и партия обновлена", variant: "default" });
        setProcessedKeys(prev => (prev.includes(selectedRow.key) ? prev : [...prev, selectedRow.key]));
        
        // Reset all form state completely
        setSelectedRow(null);
        setSelectedClient("");
        setSelectedWorkOrder("");
        setSelectedBatch("");
        
        // Reset form fields
        setClass1("");
        setClass2("");
        setClass3("");
        setRepairValue("");
        setScrapValue("");
        setStartDate("");
        setEndDate("");
        setScrapInputs({
          rattling: "",
          external: "",
          jetting: "",
          mpi: "",
          drift: "",
          emi: ""
        });
        setInitialQty(0);
        setStagesCompleted(false);

        if (sharePointService && refreshDataInBackground) {
          try {
            safeLocalStorage.removeItem("sharepoint_last_refresh");
            await refreshDataInBackground(sharePointService);
          } catch (refreshError) {
            console.warn("Failed to refresh SharePoint data after inspection save:", refreshError);
          }
        }
      } else {
        toast({ title: "Ошибка", description: "Не удалось обновить данные партии", variant: "destructive" });
      }
    } catch (error) {
      console.error("Failed to update tubing inspection data:", error);
      toast({ title: "Ошибка", description: "Не удалось обновить данные партии", variant: "destructive" });
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50">
      <Header />
      <ConfirmDialog
        open={isConfirmOpen}
        title="Save Inspection Data?"
        description="Confirm saving inspection data for this batch"
        lines={confirmLines}
        confirmText="Save"
        cancelText="Cancel"
        onConfirm={doSave}
        onCancel={() => setIsConfirmOpen(false)}
        loading={isSaving}
      />
      <div className="container mx-auto px-4 py-6">
        <div className="mb-6 flex flex-wrap items-center justify-between gap-4">
          <Button variant="ghost" onClick={() => navigate("/")} className="flex items-center gap-2 text-slate-600">
            <ArrowLeft className="h-4 w-4" />
            <span>Back to Dashboard</span>
          </Button>
          <div className="flex items-center gap-2 text-blue-600">
            <ClipboardCheck className="h-5 w-5" />
            <span>Inspection Data Entry</span>
          </div>
        </div>

        <div className="grid gap-2 lg:grid-cols-[450px_minmax(0,1fr)] items-start max-w-[1200px] mx-auto">
        <div className="lg:col-start-1 lg:row-start-1 flex flex-col gap-2 h-full">
        {/* Step 1: Batch Selection */}
        <Card className="border-2 border-blue-200 rounded-xl shadow-md flex-1">
          <CardHeader className="border-b bg-blue-50 px-4 py-3">
            <CardTitle className="text-lg font-semibold text-blue-900">Batch Selection</CardTitle>
          </CardHeader>
          <CardContent className="p-4 pt-3">
            <div className="grid gap-3 grid-cols-1 sm:grid-cols-3">
              <div className="space-y-2">
                <Label className="text-sm">Client</Label>
                <Select value={selectedClient || ""} onValueChange={value => setSelectedClient(value || "")}>
                  <SelectTrigger className="h-8 px-2 text-sm w-full">
                    <SelectValue placeholder="Choose client" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableClients.length === 0 && (
                      <div className="px-2 py-1 text-sm text-muted-foreground">No arrived batches</div>
                    )}
                    {availableClients.map(client => (
                      <SelectItem key={client} value={client}>
                        {client}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-2">
                <Label className="text-sm">Work Order</Label>
                <Select
                  value={selectedWorkOrder || ""}
                  onValueChange={value => setSelectedWorkOrder(value || "")}
                  disabled={!selectedClient}
                >
                  <SelectTrigger className="h-8 px-2 text-sm w-full">
                    <SelectValue placeholder="Choose work order" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableWorkOrders.length === 0 && (
                      <div className="px-2 py-1 text-sm text-muted-foreground">No arrived batches</div>
                    )}
                    {availableWorkOrders.map(wo => (
                      <SelectItem key={wo} value={wo}>
                        {wo}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-2">
                <Label className="text-sm">Batch</Label>
                <Select
                  value={selectedBatch || ""}
                  onValueChange={value => setSelectedBatch(value || "")}
                  disabled={!selectedClient || !selectedWorkOrder}
                >
                  <SelectTrigger className="h-8 px-2 text-sm w-full">
                    <SelectValue placeholder="Choose arrived batch" />
                  </SelectTrigger>
                  <SelectContent>
                    {availableBatches.length === 0 && (
                      <div className="px-2 py-1 text-sm text-muted-foreground">No arrived batches</div>
                    )}
                    {availableBatches.map(batch => (
                      <SelectItem key={batch.batch} value={batch.batch}>
                        {batch.batch}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </div>

            {selectedRow && (
              <div className="mt-3 flex items-start justify-between rounded-lg border border-blue-100 bg-white p-2 text-xs text-blue-900">
                <div>
                  <p className="font-semibold">Batch Info</p>
                  <p>Qty: {initialQty}</p>
                </div>
                <span className="ml-4 rounded-full bg-blue-100 px-2.5 py-0.5 text-xs font-medium uppercase tracking-wide text-blue-800">
                  {selectedRow.status || "Arrived"}
                </span>
              </div>
            )}

            {/* Qty change is intentionally not available on Inspection Data page */}
          </CardContent>
        </Card>

        {/* Step 2: Inspection Stages */}
        <Card id={stagesCardId} className={`border-2 border-blue-200 rounded-xl shadow-md flex-1 flex flex-col ${!selectedRow ? 'opacity-50' : ''}`}>
          <CardHeader className="flex flex-col gap-1 border-b bg-blue-50 px-3 py-2 sm:flex-row sm:items-center sm:justify-between">
            <CardTitle className="text-lg font-semibold text-blue-900">Inspection Stages</CardTitle>
            {selectedRow && (
              <div className="flex flex-wrap items-center gap-3 text-sm text-muted-foreground">
                <span className="rounded-full bg-blue-100 px-3 py-1 font-medium text-blue-800">
                  {selectedRow.status || "Arrived"}
                </span>
                <span>
                  Total Qty: <span className="font-semibold text-blue-900">{initialQty}</span>
                </span>
              </div>
            )}
          </CardHeader>
          <CardContent className="space-y-2 p-3 pt-2 flex-1">
            {!selectedRow && (
              <div className="rounded-lg border border-blue-200 bg-white p-2 text-center text-xs text-blue-800">
                Выберите партию в разделе "Batch Selection" для заполнения этапов инспекции
              </div>
            )}
            <div className="overflow-x-auto rounded-lg border border-blue-100">
              <Table>
                <TableHeader className="bg-blue-50 [&_th]:h-9 [&_th]:px-2.5 [&_th]:py-1.5">
                  <TableRow>
                    <TableHead className="w-1/3 text-sm font-semibold text-blue-700">Stage</TableHead>
                    <TableHead className="text-sm font-semibold text-blue-700">Qty</TableHead>
                    <TableHead className="text-sm font-semibold text-blue-700">Scrap Qty</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {stageMeta.map(stage => (
                    <TableRow key={stage.key}>
                      <TableCell className="p-2 font-medium text-slate-700">{stage.label}</TableCell>
                      <TableCell className="p-2">
                        <Input
                          value={String(computedQuantities[stage.key] ?? 0)}
                          disabled
                          className="h-8 text-sm px-2 bg-gray-100 text-gray-500 border-gray-300 cursor-not-allowed"
                        />
                      </TableCell>
                      <TableCell className="p-2">
                        {stage.scrapKey ? (
                          <Input
                            value={scrapInputs[stage.scrapKey] ?? ""}
                            onChange={e => handleScrapChange(stage.scrapKey as ScrapKey, e.target.value)}
                            inputMode="numeric"
                            placeholder="0"
                            disabled={!selectedRow}
                            className={`h-8 text-sm px-2 ${!selectedRow ? 'bg-gray-100 text-gray-500 border-gray-300 cursor-not-allowed' : ''}`}
                          />
                        ) : (
                          <span className="text-muted-foreground">—</span>
                        )}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          </CardContent>
        </Card>
        </div>

        {/* Step 3: Inspection Data */}
        <Card className={`border-emerald-100 shadow-sm lg:col-start-2 lg:row-start-1 max-w-[760px] self-start ${!stagesCompleted ? 'opacity-50' : ''}`}>
          <CardHeader className="border-b border-emerald-100 px-3 py-2">
            <CardTitle className="text-lg font-semibold text-emerald-900">Inspection Data</CardTitle>
          </CardHeader>
          <CardContent className="space-y-1 p-2 pt-2">
            {!stagesCompleted && (
              <div className="rounded-lg border border-amber-200 bg-amber-50 p-2 text-center text-xs text-amber-800">
                Заполните этапы инспекции для активации полей данных
              </div>
            )}
            <div className="grid gap-2 sm:grid-cols-2">
              <div className="space-y-2">
                <Label htmlFor="class1">Class 1</Label>
                <Input
                  id="class1"
                  value={class1}
                  onChange={event => setClass1(event.target.value)}
                  placeholder="Enter Class 1"
                  disabled={!stagesCompleted}
                  className={`h-8 text-sm ${!stagesCompleted ? 'bg-gray-100 text-gray-500 border-gray-300 cursor-not-allowed' : ''}`}
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="class2">Class 2</Label>
                <Input
                  id="class2"
                  value={class2}
                  onChange={event => setClass2(event.target.value)}
                  placeholder="Enter Class 2"
                  disabled={!stagesCompleted}
                  className={`h-8 text-sm ${!stagesCompleted ? 'bg-gray-100 text-gray-500 border-gray-300 cursor-not-allowed' : ''}`}
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="class3">Class 3</Label>
                <Input
                  id="class3"
                  value={class3}
                  onChange={event => setClass3(event.target.value)}
                  placeholder="Enter Class 3"
                  disabled={!stagesCompleted}
                  className={`h-8 text-sm ${!stagesCompleted ? 'bg-gray-100 text-gray-500 border-gray-300 cursor-not-allowed' : ''}`}
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="repair">Repair</Label>
                <Input
                  id="repair"
                  value={repairValue}
                  onChange={event => setRepairValue(sanitizeDigits(event.target.value))}
                  placeholder="0"
                  inputMode="numeric"
                  disabled={!stagesCompleted}
                  className={`h-8 text-sm ${!stagesCompleted ? 'bg-gray-100 text-gray-500 border-gray-300 cursor-not-allowed' : ''}`}
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="startDate">Start Date</Label>
                <DateInputField
                  id="startDate"
                  value={startDate}
                  onChange={setStartDate}
                  disabled={!stagesCompleted}
                  className="h-8 text-sm"
                  placeholder="dd/mm/yyyy"
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="endDate">End Date</Label>
                <DateInputField
                  id="endDate"
                  value={endDate}
                  onChange={setEndDate}
                  disabled={!stagesCompleted}
                  className="h-8 text-sm"
                  placeholder="dd/mm/yyyy"
                />
              </div>
            </div>

            {selectedRow && (
              <div className="rounded-lg border border-emerald-100 bg-emerald-50/70 p-3 text-sm text-emerald-900">
                <p className="font-semibold">Current selection</p>
                <div className="mt-1 flex flex-wrap gap-4">
                  <span>Client: {selectedRow.client}</span>
                  <span>WO: {selectedRow.wo_no}</span>
                  <span>Batch: {selectedRow.batch}</span>
                </div>
              </div>
            )}
          </CardContent>
          <CardFooter className="flex flex-wrap items-center justify-between gap-4 px-3 py-2">
            <div className="text-sm text-muted-foreground">
              Итоговый Scrap: <span className="font-semibold text-emerald-700">{totalScrap}</span>
            </div>
            <Button onClick={handleSave} disabled={isSaving || !selectedRow || !stagesCompleted} className="h-9 px-6">
              {isSaving ? "Saving..." : "Save"}
            </Button>
          </CardFooter>
        </Card>

        </div>
      </div>
    </div>
  );
}
