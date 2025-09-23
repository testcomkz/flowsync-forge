import type { ChangeEvent } from "react";
import { useEffect, useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import { format } from "date-fns";
import { ArrowLeft, Calendar as CalendarIcon, Truck } from "lucide-react";
import { Header } from "@/components/layout/Header";
import { Card, CardContent, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Input } from "@/components/ui/input";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Calendar } from "@/components/ui/calendar";
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

const parseDateInput = (value: string) => {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }

  const numericOnlyMatch = trimmed.match(/^(\d{8})$/);
  if (numericOnlyMatch) {
    const digits = numericOnlyMatch[1];
    const day = Number(digits.slice(0, 2));
    const month = Number(digits.slice(2, 4));
    const year = Number(digits.slice(4));
    const date = new Date(year, month - 1, day);
    if (date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day) {
      return format(date, "yyyy-MM-dd");
    }
    return null;
  }

  const normalized = trimmed.replace(/[.-]/g, "/").replace(/\s+/g, "");
  const parts = normalized.split("/").filter(Boolean);
  if (parts.length !== 3) {
    return null;
  }

  let year: number;
  let month: number;
  let day: number;

  if (parts[0].length === 4) {
    year = Number(parts[0]);
    month = Number(parts[1]);
    day = Number(parts[2]);
  } else if (parts[2].length === 4) {
    day = Number(parts[0]);
    month = Number(parts[1]);
    year = Number(parts[2]);
  } else {
    return null;
  }

  if (!Number.isFinite(day) || !Number.isFinite(month) || !Number.isFinite(year)) {
    return null;
  }

  if (year < 100) {
    year += year >= 70 ? 1900 : 2000;
  }

  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return null;
  }

  return format(date, "yyyy-MM-dd");
};

const toDateObject = (value: string) => {
  if (!value) return undefined;
  const [year, month, day] = value.split("-").map(Number);
  if (!year || !month || !day) return undefined;
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return undefined;
  }
  return date;
};

interface DateInputFieldProps {
  value: string;
  onChange: (value: string) => void;
  disabled?: boolean;
  placeholder?: string;
}

function DateInputField({ value, onChange, disabled, placeholder }: DateInputFieldProps) {
  const [inputValue, setInputValue] = useState(fromDateInputValue(value));
  const [isOpen, setIsOpen] = useState(false);

  useEffect(() => {
    setInputValue(fromDateInputValue(value));
  }, [value]);

  const handleInputChange = (event: ChangeEvent<HTMLInputElement>) => {
    const rawValue = event.target.value;
    setInputValue(rawValue);

    const parsed = parseDateInput(rawValue);
    if (parsed === "") {
      onChange("");
    } else if (parsed) {
      onChange(parsed);
    } else {
      onChange("");
    }
  };

  const handleBlur = () => {
    const parsed = parseDateInput(inputValue);
    if (parsed && parsed !== "") {
      setInputValue(fromDateInputValue(parsed));
      onChange(parsed);
    }
  };

  const handleSelect = (date: Date | undefined) => {
    if (!date) {
      setInputValue("");
      onChange("");
      return;
    }

    const isoValue = format(date, "yyyy-MM-dd");
    setInputValue(format(date, "dd/MM/yyyy"));
    onChange(isoValue);
    setIsOpen(false);
  };

  return (
    <div className="flex items-center gap-2">
      <Input
        value={inputValue}
        onChange={handleInputChange}
        onBlur={handleBlur}
        placeholder={placeholder}
        disabled={disabled}
        inputMode="numeric"
      />
      <Popover open={isOpen} onOpenChange={setIsOpen}>
        <PopoverTrigger asChild>
          <Button
            type="button"
            variant="outline"
            size="icon"
            className="h-10 w-10"
            disabled={disabled}
          >
            <CalendarIcon className="h-4 w-4" />
            <span className="sr-only">Выбрать дату</span>
          </Button>
        </PopoverTrigger>
        {!disabled && (
          <PopoverContent align="end" className="w-auto p-0">
            <Calendar
              mode="single"
              selected={toDateObject(value)}
              onSelect={handleSelect}
              initialFocus
            />
          </PopoverContent>
        )}
      </Popover>
    </div>
  );
}

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

  useEffect(() => {
    if (!selectedRow) {
      setLoadOutDate("");
      setActNoOper("");
      setActDate("");
      return;
    }

    setLoadOutDate(selectedRow.loadOutDate);
    setActNoOper(selectedRow.actNoOper);
    setActDate(selectedRow.actDate);
  }, [selectedRow]);

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
                      onChange={setLoadOutDate}
                      disabled={isFormDisabled || !selectedBatch}
                      placeholder="dd/mm/yyyy"
                    />
                  </div>

                  <div className="space-y-2">
                    <Label>AVR</Label>
                    <Input
                      value={actNoOper}
                      onChange={event => setActNoOper(event.target.value)}
                      placeholder="Enter AVR"
                      disabled={isFormDisabled || !selectedBatch}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label>AVR Date</Label>
                    <DateInputField
                      value={actDate}
                      onChange={setActDate}
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
