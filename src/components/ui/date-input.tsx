import type { ChangeEvent } from "react";
import { useEffect, useState } from "react";
import { format } from "date-fns";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Calendar } from "@/components/ui/calendar";
import { Calendar as CalendarIcon } from "lucide-react";

// Convert various raw values (Excel serials, strings) into ISO yyyy-mm-dd for the input value
export const toDateInputValue = (value: unknown) => {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number" && Number.isFinite(value)) {
    const excelEpoch = Date.UTC(1899, 11, 30);
    const millis = excelEpoch + value * 86400000;
    return new Date(millis).toISOString().slice(0, 10);
  }
  const stringValue = String(value).trim();
  if (!stringValue) return "";
  const isoMatch = stringValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
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
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
  return "";
};

// ISO -> display dd/MM/yyyy
const fromDateInputValue = (value: string | null | undefined) => {
  if (!value) return "";
  const isoValue = toDateInputValue(value);
  if (!isoValue) return "";
  const [year, month, day] = isoValue.split("-");
  if (!year || !month || !day) return "";
  return `${day}/${month}/${year}`;
};

// Heuristic formatter while typing: produce dd/mm/yyyy (month can be 1 or 2 digits), year up to 4 digits
function formatDisplayDateSmart(rawValue: string) {
  const src = rawValue || "";
  const normalizeDigits = (s: string) => s.replace(/\D/g, "");
  // Normalize separators to '/'
  const withSep = src.replace(/[.\-\s]+/g, "/");
  if (withSep.includes("/")) {
    const parts = withSep.split("/");
    const day = normalizeDigits(parts[0] || "").slice(0, 2);
    const monthRaw = normalizeDigits(parts[1] || "");
    let yearDigits = normalizeDigits(parts[2] || "");

    let monthDigits = monthRaw;
    if (monthRaw.length >= 2) {
      const mm2 = monthRaw.slice(0, 2);
      if (Number(mm2) >= 1 && Number(mm2) <= 12) {
        monthDigits = mm2;
        const overflow = monthRaw.slice(2);
        if (overflow) yearDigits = (yearDigits + overflow).slice(0, 4);
      } else {
        // Treat as single-digit month, push remaining digits to the END of year
        const first = monthRaw[0] || "";
        const overflow = monthRaw.slice(1);
        monthDigits = first;
        yearDigits = (yearDigits + overflow).slice(0, 4);
      }
    } else {
      monthDigits = monthRaw; // 0..1 digit while typing
    }

    yearDigits = yearDigits.slice(0, 4); // only restriction for year

    if (!day) return "";
    if (!monthDigits) return `${day}`;
    if (!yearDigits) return `${day}/${monthDigits}`;
    return `${day}/${monthDigits}/${yearDigits}`;
  }

  // Digits only path
  const digits = withSep.replace(/\D/g, "").slice(0, 8); // dd(2) + mm(<=2 or 1) + yyyy(<=4)
  if (!digits) return "";
  if (digits.length <= 2) return digits;

  const day = digits.slice(0, 2);
  let rest = digits.slice(2);
  if (rest.length === 1) return `${day}/${rest}`;

  // Prefer 2-digit month if valid (<=12), else use single-digit and push remainder to year
  let month: string;
  if (rest.length >= 2) {
    const mm2 = rest.slice(0, 2);
    if (Number(mm2) >= 1 && Number(mm2) <= 12) {
      month = mm2;
      rest = rest.slice(2);
    } else {
      month = rest.slice(0, 1);
      rest = rest.slice(1);
    }
  } else {
    month = rest;
    rest = "";
  }

  if (!rest) return `${day}/${month}`;
  const year = rest.slice(0, 4);
  return `${day}/${month}/${year}`;
}

// Parse dd/MM/yyyy (month may be 1 or 2 digits) to ISO yyyy-mm-dd; tolerate separators ' ', '.', '-', '/'
function parseDateInput(value: string) {
  const trimmed = (value || "").trim();
  if (!trimmed) return "";

  const normalized = trimmed.replace(/[.\-\s]+/g, "/");
  const parts = normalized.split("/").filter(Boolean);

  // If no explicit separators, try to reformat first
  if (parts.length === 1) {
    const formatted = formatDisplayDateSmart(trimmed);
    if (!formatted) return "";
    return parseDateInput(formatted);
  }

  if (parts.length !== 3) return null;

  const dayNum = Number(parts[0]);
  const monthNum = Number(parts[1]);
  const yearStr = parts[2].slice(0, 4); // only cap to 4 digits, no other constraints
  const yearNum = Number(yearStr);

  if (!Number.isFinite(dayNum) || !Number.isFinite(monthNum) || !Number.isFinite(yearNum)) return null;

  const date = new Date(yearNum, monthNum - 1, dayNum);
  if (date.getFullYear() !== yearNum || date.getMonth() !== monthNum - 1 || date.getDate() !== dayNum) return null;
  return format(date, "yyyy-MM-dd");
}

function toDateObject(value: string) {
  if (!value) return undefined;
  const [y, m, d] = value.split("-").map(Number);
  if (!y || !m || !d) return undefined;
  const date = new Date(y, m - 1, d);
  if (date.getFullYear() !== y || date.getMonth() !== m - 1 || date.getDate() !== d) return undefined;
  return date;
}

export interface DateInputFieldProps {
  id?: string;
  value: string; // ISO yyyy-mm-dd or empty
  onChange: (value: string) => void; // emits ISO or ""
  disabled?: boolean;
  placeholder?: string;
  className?: string; // optional extra classes for Input to fit compact rows
}

export function DateInputField({ id, value, onChange, disabled, placeholder, className }: DateInputFieldProps) {
  const [inputValue, setInputValue] = useState(fromDateInputValue(value));
  const [isOpen, setIsOpen] = useState(false);

  useEffect(() => {
    setInputValue(fromDateInputValue(value));
  }, [value]);

  const handleInputChange = (event: ChangeEvent<HTMLInputElement>) => {
    const rawValue = event.target.value;
    const formattedForDisplay = formatDisplayDateSmart(rawValue);
    setInputValue(formattedForDisplay);

    const parsed = parseDateInput(formattedForDisplay);
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
      setInputValue(fromDateInputValue(parsed)); // dd/MM/yyyy with zero-padded month/day
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
        id={id}
        value={inputValue}
        onChange={handleInputChange}
        onBlur={handleBlur}
        placeholder={placeholder}
        disabled={disabled}
        inputMode="numeric"
        className={`h-10 flex-1 rounded-xl border-sky-200 bg-white/90 text-sky-900 shadow-sm transition focus-visible:border-sky-400 focus-visible:ring-sky-200 disabled:opacity-70 ${className ? className : ""}`}
      />
      <Popover open={isOpen} onOpenChange={setIsOpen}>
        <PopoverTrigger asChild>
          <Button
            type="button"
            variant="outline"
            size="icon"
            className="h-10 w-10 rounded-xl border-sky-200 bg-white/90 text-sky-500 shadow-sm transition hover:bg-sky-50 focus-visible:ring-sky-200"
            disabled={disabled}
          >
            <CalendarIcon className="h-4 w-4" />
            <span className="sr-only">Choose date</span>
          </Button>
        </PopoverTrigger>
        {!disabled && (
          <PopoverContent align="end" className="w-auto rounded-xl border border-sky-100 bg-white p-2 shadow-lg">
            <Calendar mode="single" selected={toDateObject(value)} onSelect={handleSelect} initialFocus />
          </PopoverContent>
        )}
      </Popover>
    </div>
  );
}
