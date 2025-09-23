import { useCallback, useMemo, type ChangeEvent } from "react";
import { cn } from "@/lib/utils";

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

export interface DateInputFieldProps {
  id?: string;
  label?: string;
  value?: string; // ISO yyyy-mm-dd or empty
  onChange: (value: string) => void; // emits ISO or ""
  disabled?: boolean;
  placeholder?: string;
  className?: string; // optional extra classes for layout tweaks
}

const isoToDate = (value: string | undefined): Date | null => {
  const isoValue = toDateInputValue(value);
  if (!isoValue) return null;
  const parts = isoValue.split("-").map(Number);
  if (parts.length !== 3) return null;
  const [y, m, d] = parts;
  const dt = new Date(y, (m || 1) - 1, d || 1);
  if (
    dt.getFullYear() !== y ||
    dt.getMonth() !== (m - 1) ||
    dt.getDate() !== d
  ) {
    return null;
  }
  return dt;
};

const dateToIso = (value: Date | null): string => {
  if (!value) return "";
  const year = String(value.getFullYear()).padStart(4, "0");
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const day = String(value.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};

const formatDisplay = (value: Date | null): string => {
  if (!value) return "";
  const dd = String(value.getDate()).padStart(2, "0");
  const mm = String(value.getMonth() + 1).padStart(2, "0");
  const yyyy = String(value.getFullYear());
  return `${dd}/${mm}/${yyyy}`;
};

const parseDdMmYyyy = (text: string): Date | null => {
  const m = text.match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{4})\s*$/);
  if (!m) return null;
  const dd = Number(m[1]);
  const mm = Number(m[2]);
  const yyyy = Number(m[3]);
  if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return null;
  const dt = new Date(yyyy, mm - 1, dd);
  if (
    dt.getFullYear() !== yyyy ||
    dt.getMonth() !== (mm - 1) ||
    dt.getDate() !== dd
  ) {
    return null;
  }
  return dt;
};

export function DateInputField({
  id,
  label,
  value,
  onChange,
  disabled,
  placeholder,
  className,
}: DateInputFieldProps) {
  const normalizedValue = useMemo(() => toDateInputValue(value), [value]);

  const handleChange = useCallback(
    (e: ChangeEvent<HTMLInputElement>) => {
      // Native date input emits ISO yyyy-mm-dd or empty
      onChange(toDateInputValue(e.target.value));
    },
    [onChange],
  );

  const ariaLabel = label ? undefined : !id ? placeholder ?? "Select date" : undefined;

  return (
    <div className="w-full">
      {label ? (
        <label htmlFor={id} className="mb-1 block text-sm font-semibold text-slate-700">
          {label}
        </label>
      ) : null}
      <input
        id={id}
        type="date"
        value={normalizedValue}
        onChange={handleChange}
        placeholder={placeholder ?? "dd/mm/yyyy"}
        aria-label={ariaLabel}
        disabled={disabled}
        className={cn(
          "h-11 w-full rounded-md border border-gray-300 bg-white px-3 text-gray-900 shadow-sm focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500 focus-visible:border-blue-500 disabled:cursor-not-allowed disabled:bg-gray-100 disabled:text-gray-500 disabled:border-gray-300",
          className,
        )}
      />
    </div>
  );
}
