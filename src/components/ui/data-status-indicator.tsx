import { Badge } from "@/components/ui/badge";
import { CheckCircle, Clock, Database, Wifi, WifiOff } from "lucide-react";

interface DataStatusIndicatorProps {
  isConnected: boolean;
  isLoading: boolean;
  lastUpdate?: string;
  dataCount?: number;
}

export function DataStatusIndicator({ 
  isConnected, 
  isLoading, 
  lastUpdate, 
  dataCount 
}: DataStatusIndicatorProps) {
  const getStatusInfo = () => {
    if (isLoading) {
      return {
        icon: Clock,
        text: "Updating...",
        variant: "secondary" as const,
        className: "animate-pulse"
      };
    }
    
    if (!isConnected) {
      return {
        icon: WifiOff,
        text: "Not Connected",
        variant: "destructive" as const,
        className: ""
      };
    }
    
    return {
      icon: CheckCircle,
      text: "Connected",
      variant: "default" as const,
      className: "bg-green-100 text-green-800 border-green-300"
    };
  };

  const status = getStatusInfo();
  const Icon = status.icon;
  
  const formatLastUpdate = (timestamp?: string) => {
    if (!timestamp) return "";
    
    try {
      const date = new Date(timestamp);
      const now = new Date();
      const diffMs = now.getTime() - date.getTime();
      const diffMinutes = Math.floor(diffMs / (1000 * 60));
      
      if (diffMinutes < 1) return "Just now";
      if (diffMinutes < 60) return `${diffMinutes}m ago`;
      
      const diffHours = Math.floor(diffMinutes / 60);
      if (diffHours < 24) return `${diffHours}h ago`;
      
      return date.toLocaleDateString();
    } catch {
      return "";
    }
  };

  return (
    <div className="flex items-center space-x-2 text-sm text-muted-foreground">
      <Badge variant={status.variant} className={`${status.className} flex items-center space-x-1`}>
        <Icon className="h-3 w-3" />
        <span>{status.text}</span>
      </Badge>
      
      {dataCount !== undefined && (
        <div className="flex items-center space-x-1">
          <Database className="h-3 w-3" />
          <span>{dataCount.toLocaleString()} records</span>
        </div>
      )}
      
      {lastUpdate && (
        <div className="flex items-center space-x-1">
          <Clock className="h-3 w-3" />
          <span>{formatLastUpdate(lastUpdate)}</span>
        </div>
      )}
    </div>
  );
}
