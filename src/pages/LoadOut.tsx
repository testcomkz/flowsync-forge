import { useEffect, useMemo, useRef, useState } from "react";
import { useNavigate } from "react-router-dom";
import { ArrowLeft, Truck } from "lucide-react";
import { Header } from "@/components/layout/Header";
import { Card, CardContent, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Input } from "@/components/ui/input";
import { DateInputField } from "@/components/ui/date-input";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useToast } from "@/hooks/use-toast";
import { useSharePointInstantData } from "@/hooks/useInstantData";
import { useAuth } from "@/contexts/AuthContext";

interface LoadOutRow {
  client: string;
  wo_no: string;
  batch: string;
  status: string;
  loadOutDate: string;
  actNoOper: string;
  actDate: string;
}

const normalizeValue = (value: unknown) =>
  value === null || value === undefined ? "" : String(value).trim();

const canonicalize = (value: unknown) =>
  normalizeValue(value)
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[_-]+/g, "")
    .replace(/[^a-z0-9]/g, "");

const toDateInputValue = (value: unknown) => {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    const excelEpoch = Date.UTC(1899, 11, 30);
    const millis = excelEpoch + value * 86400000;
    return new Date(millis).toISOString().slice(0, 10);
  }

  const stringValue = normalizeValue(value);
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

const fromDateInputValue = (value: string | null | undefined) => {
  if (!value) return "";
  const isoValue = toDateInputValue(value);
  if (!isoValue) return "";
  const [year, month, day] = isoValue.split("-");
  if (!year || !month || !day) return "";
  return `${day}/${month}/${year}`;
};

// Date input is provided by shared component `@/components/ui/date-input` now

export default function LoadOut() {
  const navigate = useNavigate();
  const { user } = useAuth();
  const { toast } = useToast();
  const { tubingData } = useSharePointInstantData();
  const { sharePointService, isConnected, refreshDataInBackground } = useSharePoint();

  const [selectedClient, setSelectedClient] = useState("");
  const [selectedWorkOrder, setSelectedWorkOrder] = useState("");
  const [selectedBatch, setSelectedBatch] = useState("");
  const [loadOutDate, setLoadOutDate] = useState("");
  const [actNoOper, setActNoOper] = useState("");
  const [actDate, setActDate] = useState("");
  const [isLoadOutDateDirty, setIsLoadOutDateDirty] = useState(false);
  const [isActNoOperDirty, setIsActNoOperDirty] = useState(false);
  const [isActDateDirty, setIsActDateDirty] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  const loadOutRows = useMemo(() => {
    if (!Array.isArray(tubingData) || tubingData.length < 2) {
      return [] as LoadOutRow[];
    }

    const headersRow = tubingData[0];
    if (!Array.isArray(headersRow)) {
      return [] as LoadOutRow[];
    }

    const findIndex = (matcher: (header: string) => boolean) =>
      headersRow.findIndex(header => matcher(canonicalize(header)));

    const clientIndex = findIndex(header => header.includes("client"));
    const woIndex = findIndex(header => header.includes("wo"));
    const batchIndex = findIndex(header => header.includes("batch"));
    const statusIndex = findIndex(header => header.includes("status"));
    const loadOutDateIndex = findIndex(header => header.includes("loadoutdate"));
    const actNoOperIndex = findIndex(header => header.includes("actnooper"));
    const actDateIndex = findIndex(header => header.includes("actdate"));

    return tubingData.slice(1).reduce<LoadOutRow[]>((acc, rowRaw) => {
      if (!Array.isArray(rowRaw)) return acc;

      const row = rowRaw as unknown[];
      const client = normalizeValue(row[clientIndex]);
      const wo_no = normalizeValue(row[woIndex]);
      const batch = normalizeValue(row[batchIndex]);
      const status = normalizeValue(row[statusIndex]);
      const statusKey = canonicalize(status);

      if (!client || !wo_no || !batch) return acc;
      if (!statusKey.includes("inspectiondone")) return acc;

      acc.push({
        client,
        wo_no,
        batch,
        status,
        loadOutDate: toDateInputValue(row[loadOutDateIndex]),
        actNoOper: normalizeValue(row[actNoOperIndex]),
        actDate: toDateInputValue(row[actDateIndex]),
      });
      return acc;
    }, []);
  }, [tubingData]);

  const availableClients = useMemo(() => {
    const unique = new Set<string>();
    loadOutRows.forEach(row => {
      if (row.client) unique.add(row.client);
    });
    return Array.from(unique).sort((a, b) => a.localeCompare(b));
  }, [loadOutRows]);

  const availableWorkOrders = useMemo(() => {
    const unique = new Set<string>();
    loadOutRows.forEach(row => {
      if (row.client === selectedClient && row.wo_no) {
        unique.add(row.wo_no);
      }
    });
    return Array.from(unique).sort((a, b) => a.localeCompare(b));
  }, [loadOutRows, selectedClient]);

  const availableBatches = useMemo(() => {
    const unique = new Set<string>();
    loadOutRows
      .filter(row => row.client === selectedClient && row.wo_no === selectedWorkOrder)
      .forEach(row => {
        if (row.batch) unique.add(row.batch);
      });
    return Array.from(unique).sort((a, b) => a.localeCompare(b));
  }, [loadOutRows, selectedClient, selectedWorkOrder]);

  const selectedRow = useMemo(
    () =>
      loadOutRows.find(
        row =>
          row.client === selectedClient &&
          row.wo_no === selectedWorkOrder &&
          row.batch === selectedBatch
      ) || null,
    [loadOutRows, selectedClient, selectedWorkOrder, selectedBatch]
  );

  const selectedRowKey = selectedRow ? `${selectedRow.client}|${selectedRow.wo_no}|${selectedRow.batch}` : "";
  const previousRowKeyRef = useRef<string>("");

  useEffect(() => {
    if (!selectedRow) {
      setLoadOutDate("");
      setActNoOper("");
      setActDate("");
      setIsLoadOutDateDirty(false);
      setIsActNoOperDirty(false);
      setIsActDateDirty(false);
      previousRowKeyRef.current = "";
      return;
    }

    const isNewRow = previousRowKeyRef.current !== selectedRowKey;

    if (!isLoadOutDateDirty || isNewRow) {
      setLoadOutDate(selectedRow.loadOutDate);
    }
    if (!isActNoOperDirty || isNewRow) {
      setActNoOper(selectedRow.actNoOper);
    }
    if (!isActDateDirty || isNewRow) {
      setActDate(selectedRow.actDate);
    }

    if (isNewRow) {
      setIsLoadOutDateDirty(false);
      setIsActNoOperDirty(false);
      setIsActDateDirty(false);
    }

    previousRowKeyRef.current = selectedRowKey;
  }, [
    selectedRowKey,
    selectedRow,
    isLoadOutDateDirty,
    isActNoOperDirty,
    isActDateDirty,
    selectedRow?.loadOutDate,
    selectedRow?.actNoOper,
    selectedRow?.actDate,
  ]);

  const handleLoadOutDateChange = (next: string) => {
    setLoadOutDate(next);
    setIsLoadOutDateDirty(next !== (selectedRow?.loadOutDate ?? ""));
  };

  const handleActDateChange = (next: string) => {
    setActDate(next);
    setIsActDateDirty(next !== (selectedRow?.actDate ?? ""));
  };

  const handleActNoOperChange = (next: string) => {
    setActNoOper(next);
    setIsActNoOperDirty(next !== (selectedRow?.actNoOper ?? ""));
  };

  const handleSave = async (event: React.FormEvent) => {
    event.preventDefault();

    if (!user) {
      toast({
        title: "Ошибка",
        description: "Пожалуйста, войдите в систему",
        variant: "destructive",
      });
      return;
    }

    if (!sharePointService || !isConnected) {
      toast({
        title: "SharePoint не подключен",
        description: "Подключитесь к SharePoint перед сохранением",
        variant: "destructive",
      });
      return;
    }

    if (!selectedClient || !selectedWorkOrder || !selectedBatch || !loadOutDate || !actNoOper || !actDate) {
      toast({
        title: "Заполните все поля",
        description: "Выберите партию и заполните Load Out Date, AVR и AVR Date",
        variant: "destructive",
      });
      return;
    }

    setIsSaving(true);
    try {
      const success = await sharePointService.updateLoadOutData({
        client: selectedClient,
        wo_no: selectedWorkOrder,
        batch: selectedBatch,
        load_out_date: loadOutDate,
        act_no_oper: actNoOper,
        act_date: actDate,
        status: "Completed",
      });

      if (success) {
        toast({
          title: "Load Out сохранен",
          description: `${selectedBatch} обновлен и помечен как Completed`,
        });

        if (sharePointService && refreshDataInBackground) {
          try {
            localStorage.removeItem("sharepoint_last_refresh");
            await refreshDataInBackground(sharePointService);
          } catch (refreshError) {
            console.warn("Failed to refresh SharePoint data after load out save:", refreshError);
          }
        }

        setSelectedBatch("");
      } else {
        toast({
          title: "Ошибка",
          description: "Не удалось обновить данные Load Out",
          variant: "destructive",
        });
      }
    } catch (error) {
      console.error("Failed to save load out data:", error);
      toast({
        title: "Ошибка",
        description: "Не удалось обновить данные Load Out",
        variant: "destructive",
      });
    } finally {
      setIsSaving(false);
    }
  };

  const isFormDisabled = !isConnected || loadOutRows.length === 0;
  const canSave =
    !isFormDisabled &&
    !!selectedClient &&
    !!selectedWorkOrder &&
    !!selectedBatch &&
    !!loadOutDate &&
    !!actNoOper &&
    !!actDate;

  return (
    <div className="min-h-screen bg-gray-50">
      <Header />
      <div className="container mx-auto px-4 py-6">
        <div className="mb-6 flex flex-wrap items-center justify-between gap-4">
          <Button variant="outline" onClick={() => navigate("/")} className="flex items-center gap-2">
            <ArrowLeft className="h-4 w-4" />
            <span>Back to Dashboard</span>
          </Button>
          <div className="flex items-center gap-2 text-gray-600">
            <Truck className="h-5 w-5" />
            <span>Load Out</span>
          </div>
        </div>

        <form onSubmit={handleSave}>
          <Card className="mx-auto max-w-3xl border-sky-100 shadow-sm">
            <CardHeader className="border-b border-sky-100 pb-4">
              <CardTitle className="text-lg font-semibold text-sky-900">Finalize Batch Load Out</CardTitle>
            </CardHeader>
            <CardContent className="space-y-5 pt-4">
              {!isConnected && (
                <div className="rounded-md border border-amber-200 bg-amber-50 p-3 text-sm text-amber-800">
                  Подключитесь к SharePoint, чтобы сохранять изменения.
                </div>
              )}

              {loadOutRows.length === 0 ? (
                <div className="rounded-md border border-muted bg-muted/20 p-4 text-center text-sm text-muted-foreground">
                  Нет партий со статусом Inspection Done, доступных для Load Out.
                </div>
              ) : (
                <div className="grid gap-4 sm:grid-cols-2">
                  <div className="space-y-2">
                    <Label>Client</Label>
                    <Select
                      value={selectedClient || undefined}
                      onValueChange={value => {
                        setSelectedClient(value);
                        setSelectedWorkOrder("");
                        setSelectedBatch("");
                      }}
                      disabled={isFormDisabled}
                    >
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
                      value={selectedWorkOrder || undefined}
                      onValueChange={value => {
                        setSelectedWorkOrder(value);
                        setSelectedBatch("");
                      }}
                      disabled={isFormDisabled || availableWorkOrders.length === 0}
                    >
                      <SelectTrigger>
                        <SelectValue placeholder="Select work order" />
                      </SelectTrigger>
                      <SelectContent>
                        {availableWorkOrders.map(wo => (
                          <SelectItem key={wo} value={wo}>
                            {wo}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>

                  <div className="space-y-2">
                    <Label>Batch</Label>
                    <Select
                      value={selectedBatch || undefined}
                      onValueChange={value => setSelectedBatch(value)}
                      disabled={
                        isFormDisabled ||
                        availableBatches.length === 0 ||
                        !selectedClient ||
                        !selectedWorkOrder
                      }
                    >
                      <SelectTrigger>
                        <SelectValue placeholder="Select batch" />
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

                  <div className="space-y-2">
                    <Label>Load Out Date</Label>
                    <DateInputField
                      value={loadOutDate}
                      onChange={handleLoadOutDateChange}
                      disabled={isFormDisabled || !selectedBatch}
                      placeholder="dd/mm/yyyy"
                    />
                  </div>

                  <div className="space-y-2">
                    <Label>AVR</Label>
                    <Input
                      value={actNoOper}
                      onChange={event => handleActNoOperChange(event.target.value)}
                      placeholder="Enter AVR"
                      disabled={isFormDisabled || !selectedBatch}
                      className="h-10 rounded-xl border-sky-200 bg-white/90 text-sky-900 shadow-sm transition focus-visible:border-sky-400 focus-visible:ring-sky-200 disabled:opacity-70"
                    />
                  </div>

                  <div className="space-y-2">
                    <Label>AVR Date</Label>
                    <DateInputField
                      value={actDate}
                      onChange={handleActDateChange}
                      disabled={isFormDisabled || !selectedBatch}
                      placeholder="dd/mm/yyyy"
                    />
                  </div>
                </div>
              )}
            </CardContent>
            <CardFooter className="flex items-center justify-between border-t border-sky-100 bg-sky-50/40 py-4">
              <div className="text-sm text-muted-foreground">
                Выбранная партия будет переведена в статус Completed.
              </div>
              <Button type="submit" disabled={!canSave || isSaving}>
                {isSaving ? "Saving..." : "Save"}
              </Button>
            </CardFooter>
          </Card>
        </form>
      </div>
    </div>
  );
}
