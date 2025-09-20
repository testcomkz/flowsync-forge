import { DashboardLayout } from "@/components/layout/dashboard-layout";
import { StatsCard } from "@/components/ui/stats-card";
import { BoardTable } from "@/components/dashboard/board-table";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { 
  FileText, 
  Package, 
  ClipboardCheck, 
  CheckCircle, 
  Clock,
  Plus,
  BarChart3
} from "lucide-react";
import { BoardItem, DashboardStats } from "@/types";
import { useState } from "react";

// Mock data - replace with real data from Supabase
const mockStats: DashboardStats = {
  totalRecords2025: 1247,
  awaitingInspection: 23,
  inspectionDone: 156,
  awaitingAVR: 12,
  completed: 1056,
};

const mockBoardData: BoardItem[] = [
  {
    id: "1",
    rack: "R-001",
    client: "ACME Corp",
    batch: "Batch # 001",
    wo: "WO-2025-000001",
    qty: 150,
    status: "RECEIVED, AWAITING INSPECTION",
    order_key: "WO-2025-000001-000001",
    registered_at: "2025-01-15T10:30:00Z",
  },
  {
    id: "2",
    rack: "R-002",
    client: "Beta Industries",
    batch: "Batch # 002",
    wo: "WO-2025-000002",
    qty: 200,
    status: "INSPECTION DONE",
    order_key: "WO-2025-000002-000002",
    registered_at: "2025-01-14T14:20:00Z",
  },
  {
    id: "3",
    rack: "R-003",
    client: "Gamma LLC",
    batch: "Batch # 003",
    wo: "WO-2025-000003",
    qty: 75,
    status: "AWAITING AVR",
    order_key: "WO-2025-000003-000003",
    registered_at: "2025-01-13T09:15:00Z",
  },
];

export default function Dashboard() {
  const [selectedItem, setSelectedItem] = useState<BoardItem | null>(null);

  const handleEdit = (item: BoardItem) => {
    setSelectedItem(item);
    // Open edit modal - implement later
    console.log("Edit item:", item);
  };

  return (
    <DashboardLayout currentPage="Dashboard">
      <div className="space-y-6">
        {/* Page Header */}
        <div className="flex justify-between items-center">
          <div>
            <h1 className="text-3xl font-bold tracking-tight">Dashboard</h1>
            <p className="text-muted-foreground">
              Monitor your work orders and batch processing pipeline
            </p>
          </div>
          <div className="flex space-x-2">
            <Button variant="outline">
              <BarChart3 className="h-4 w-4 mr-2" />
              Analytics
            </Button>
            <Button>
              <Plus className="h-4 w-4 mr-2" />
              New Work Order
            </Button>
          </div>
        </div>

        {/* Stats Cards */}
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-4">
          <StatsCard
            title="Total Records (2025)"
            value={mockStats.totalRecords2025.toLocaleString()}
            icon={FileText}
            variant="primary"
            trend={{ value: 12, positive: true }}
          />
          <StatsCard
            title="Awaiting Inspection"
            value={mockStats.awaitingInspection}
            icon={Clock}
            variant="warning"
            description="Needs attention"
          />
          <StatsCard
            title="Inspection Done"
            value={mockStats.inspectionDone}
            icon={ClipboardCheck}
            variant="success"
          />
          <StatsCard
            title="Awaiting AVR"
            value={mockStats.awaitingAVR}
            icon={Package}
            variant="default"
          />
          <StatsCard
            title="Completed"
            value={mockStats.completed}
            icon={CheckCircle}
            variant="success"
            trend={{ value: 8, positive: true }}
          />
        </div>

        {/* Quick Actions */}
        <Card>
          <CardHeader>
            <CardTitle>Quick Actions</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
              <Button variant="outline" className="h-20 flex-col">
                <Package className="h-6 w-6 mb-2" />
                <span className="text-xs">Create WO</span>
              </Button>
              <Button variant="outline" className="h-20 flex-col">
                <ClipboardCheck className="h-6 w-6 mb-2" />
                <span className="text-xs">Batch Registry</span>
              </Button>
              <Button variant="outline" className="h-20 flex-col">
                <CheckCircle className="h-6 w-6 mb-2" />
                <span className="text-xs">Inspection</span>
              </Button>
              <Button variant="outline" className="h-20 flex-col">
                <FileText className="h-6 w-6 mb-2" />
                <span className="text-xs">Load Out</span>
              </Button>
              <Button variant="outline" className="h-20 flex-col">
                <FileText className="h-6 w-6 mb-2" />
                <span className="text-xs">Create AVR</span>
              </Button>
              <Button variant="outline" className="h-20 flex-col">
                <BarChart3 className="h-6 w-6 mb-2" />
                <span className="text-xs">Reports</span>
              </Button>
            </div>
          </CardContent>
        </Card>

        {/* Board Table */}
        <Card>
          <CardHeader>
            <CardTitle>Board - Current Batches</CardTitle>
          </CardHeader>
          <CardContent>
            <BoardTable data={mockBoardData} onEdit={handleEdit} />
          </CardContent>
        </Card>
      </div>
    </DashboardLayout>
  );
}