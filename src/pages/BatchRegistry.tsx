import { DashboardLayout } from "@/components/layout/dashboard-layout";
import { BatchRegistryForm } from "@/components/forms/batch-registry-form";
import { BatchRegistryFormData } from "@/types";
import { useToast } from "@/hooks/use-toast";

// Mock work orders - replace with real data from Supabase
const mockWorkOrders = [
  { id: "1", wo_num: 1001, client: "ACME Corp" },
  { id: "2", wo_num: 1002, client: "Beta Industries" },
  { id: "3", wo_num: 1003, client: "Gamma LLC" },
];

export default function BatchRegistry() {
  const { toast } = useToast();

  const handleSubmit = (data: BatchRegistryFormData) => {
    // TODO: Replace with actual Supabase integration
    console.log("Registering Batch:", data);
    
    // Simulate API call
    setTimeout(() => {
      toast({
        title: "Batch Registered",
        description: `Batch has been registered successfully with status "RECEIVED, AWAITING INSPECTION".`,
      });
    }, 1000);
  };

  return (
    <DashboardLayout currentPage="Batch Registry">
      <div className="space-y-6">
        <div>
          <h1 className="text-3xl font-bold tracking-tight">Batch Registry</h1>
          <p className="text-muted-foreground">
            Register a new batch for an existing work order
          </p>
        </div>

        <BatchRegistryForm 
          onSubmit={handleSubmit} 
          workOrders={mockWorkOrders}
        />
      </div>
    </DashboardLayout>
  );
}