import { useCallback, useMemo } from "react";
import { DatePicker } from "@heroui/react";
import { parseDate, type CalendarDate } from "@internationalized/date";

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

const isoToCalendarDate = (value: string | undefined): CalendarDate | null => {
  const isoValue = toDateInputValue(value);
  if (!isoValue) return null;

  try {
    return parseDate(isoValue);
  } catch {
    return null;
  }
};

const calendarDateToIso = (value: CalendarDate | null): string => {
  if (!value) return "";

  const year = String(value.year).padStart(4, "0");
  const month = String(value.month).padStart(2, "0");
  const day = String(value.day).padStart(2, "0");

  return `${year}-${month}-${day}`;
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
  const selectedDate = useMemo(() => isoToCalendarDate(value), [value]);

  const handleChange = useCallback(
    (next: CalendarDate | null) => {
      onChange(calendarDateToIso(next));
    },
    [onChange],
  );

  return (
    <DatePicker<CalendarDate>
      id={id}
      value={selectedDate}
      onChange={handleChange}
      label={label}
      labelPlacement={label ? "outside" : undefined}
      aria-label={label ? undefined : !id ? placeholder ?? "Select date" : undefined}
      isDisabled={disabled}
      locale="en-GB"
      granularity="day"
      shouldForceLeadingZeros
      selectorButtonPlacement="end"
      className={cn("w-full", className)}
      classNames={{
        base: "w-full",
        label: "text-sm font-semibold text-slate-700",
        inputWrapper:
          "h-10 w-full rounded-xl border border-sky-200 bg-white/95 text-sky-900 shadow-sm transition data-[hover=true]:border-sky-300 data-[focus-visible=true]:ring-2 data-[focus-visible=true]:ring-sky-200 data-[focus-visible=true]:ring-offset-2",
        segment: "text-sky-900",
        selectorButton:
          "text-sky-500 data-[focus-visible=true]:outline-none data-[focus-visible=true]:ring-2 data-[focus-visible=true]:ring-sky-200",
        selectorIcon: "text-sky-500",
        popoverContent: "rounded-xl border border-sky-100 bg-white shadow-xl",
        calendar: "rounded-xl border border-sky-100 bg-white",
        calendarContent: "rounded-lg bg-white",
      }}
      popoverProps={{ placement: "bottom-end", offset: 10 }}
      calendarProps={{
        showMonthAndYearPickers: true,
        weekdayStyle: "short",
        classNames: {
          base: "rounded-xl",
          headerWrapper: "rounded-t-xl",
        },
      }}
    />
  );
}
