import { cn } from "@/lib/utils";
import { cva, type VariantProps } from "class-variance-authority";

const statusBadgeVariants = cva(
  "inline-flex items-center rounded-full px-2.5 py-0.5 text-xs font-medium transition-colors",
  {
    variants: {
      status: {
        open: "bg-status-open-bg text-status-open",
        received: "bg-status-received-bg text-status-received",
        inspection: "bg-status-inspection-bg text-status-inspection",
        done: "bg-status-done-bg text-status-done",
        awaiting: "bg-primary-light text-primary",
      },
    },
    defaultVariants: {
      status: "open",
    },
  }
);

export interface StatusBadgeProps
  extends React.HTMLAttributes<HTMLDivElement>,
    VariantProps<typeof statusBadgeVariants> {
  status: "open" | "received" | "inspection" | "done" | "awaiting";
}

function StatusBadge({ className, status, ...props }: StatusBadgeProps) {
  const getStatusText = (status: string) => {
    switch (status) {
      case "open":
        return "Open";
      case "received":
        return "Received";
      case "inspection":
        return "Inspection Done";
      case "done":
        return "Completed";
      case "awaiting":
        return "Awaiting";
      default:
        return status;
    }
  };

  return (
    <div className={cn(statusBadgeVariants({ status }), className)} {...props}>
      {getStatusText(status)}
    </div>
  );
}

export { StatusBadge, statusBadgeVariants };