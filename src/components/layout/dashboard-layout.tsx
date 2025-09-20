import { ReactNode } from "react";
import { Button } from "@/components/ui/button";
import { 
  FileText, 
  Package, 
  ClipboardCheck, 
  Truck, 
  FileCheck, 
  Settings,
  LogOut,
  User,
  Download,
  RefreshCw
} from "lucide-react";
import { useSharePoint } from "@/contexts/SharePointContext";
import { DataStatusIndicator } from "@/components/ui/data-status-indicator";
import { cn } from "@/lib/utils";

interface DashboardLayoutProps {
  children: ReactNode;
  currentPage?: string;
}

const navigationItems = [
  { name: "Dashboard", href: "/", icon: FileText },
  { name: "Create WO", href: "/create-wo", icon: Package },
  { name: "Batch Registry", href: "/batch-registry", icon: ClipboardCheck },
  { name: "Complete Inspection", href: "/inspection", icon: FileCheck },
  { name: "Create Load Out", href: "/load-out", icon: Truck },
  { name: "Create AVR", href: "/avr", icon: FileCheck },
];

export function DashboardLayout({ children, currentPage = "Dashboard" }: DashboardLayoutProps) {
  const { 
    isConnected, 
    isConnecting, 
    connectToSharePoint, 
    refreshDataInBackground, 
    sharePointService,
    isDataLoading,
    cachedClients,
    cachedWorkOrders
  } = useSharePoint();

  const handleLoadData = async () => {
    if (!isConnected) {
      await connectToSharePoint();
    }
    if (sharePointService) {
      await refreshDataInBackground(sharePointService);
    }
  };

  const handleUpdateData = async () => {
    if (sharePointService) {
      // Принудительное обновление - очищаем время последнего обновления
      localStorage.removeItem('sharepoint_last_refresh');
      await refreshDataInBackground(sharePointService);
    }
  };

  return (
    <div className="min-h-screen bg-background">
      {/* Header */}
      <header className="border-b bg-card/50 backdrop-blur supports-[backdrop-filter]:bg-card/50">
        <div className="flex h-16 items-center px-6">
          <div className="flex items-center space-x-4">
            <div className="flex items-center space-x-2">
              <div className="h-8 w-8 rounded bg-gradient-to-br from-primary to-primary-hover flex items-center justify-center">
                <Package className="h-4 w-4 text-primary-foreground" />
              </div>
              <h1 className="text-xl font-semibold">FlowSync Forge</h1>
            </div>
          </div>
          
          <div className="ml-auto flex items-center space-x-2">
            {/* Data Control Buttons */}
            {!isConnected ? (
              <Button 
                onClick={handleLoadData}
                disabled={isConnecting}
                className="bg-green-600 hover:bg-green-700"
              >
                <Download className="h-4 w-4 mr-2" />
                {isConnecting ? "Loading..." : "Load Data"}
              </Button>
            ) : (
              <Button 
                onClick={handleUpdateData}
                disabled={isDataLoading}
                variant="outline"
                className="border-blue-500 text-blue-600 hover:bg-blue-50"
              >
                <RefreshCw className={`h-4 w-4 mr-2 ${isDataLoading ? 'animate-spin' : ''}`} />
                {isDataLoading ? "Updating..." : "Update Data"}
              </Button>
            )}
            
            {/* Data Status Indicator */}
            <DataStatusIndicator 
              isConnected={isConnected}
              isLoading={isDataLoading}
              lastUpdate={localStorage.getItem('sharepoint_last_refresh') || undefined}
              dataCount={cachedClients.length + cachedWorkOrders.length}
            />
            
            <Button variant="ghost" size="sm">
              <User className="h-4 w-4 mr-2" />
              User Profile
            </Button>
            <Button variant="ghost" size="sm">
              <Settings className="h-4 w-4 mr-2" />
              Settings
            </Button>
            <Button variant="ghost" size="sm">
              <LogOut className="h-4 w-4 mr-2" />
              Logout
            </Button>
          </div>
        </div>
      </header>

      <div className="flex">
        {/* Sidebar */}
        <aside className="w-64 border-r bg-card/30 h-[calc(100vh-4rem)]">
          <nav className="p-4 space-y-2">
            {navigationItems.map((item) => {
              const Icon = item.icon;
              const isActive = currentPage === item.name;
              
              return (
                <Button
                  key={item.name}
                  variant={isActive ? "secondary" : "ghost"}
                  className={cn(
                    "w-full justify-start",
                    isActive && "bg-primary/10 text-primary hover:bg-primary/15"
                  )}
                  asChild
                >
                  <a href={item.href}>
                    <Icon className="h-4 w-4 mr-3" />
                    {item.name}
                  </a>
                </Button>
              );
            })}
          </nav>
        </aside>

        {/* Main Content */}
        <main className="flex-1 p-6">
          {children}
        </main>
      </div>
    </div>
  );
}