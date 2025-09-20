import { DashboardLayout } from "@/components/layout/dashboard-layout";
import { CreateWOForm } from "@/components/forms/create-wo-form";

import { useToast } from "@/hooks/use-toast";

export default function CreateWO() {
  const { toast } = useToast();

  const handleSubmit = (data: any) => {
    // TODO: Replace with actual Supabase integration
    console.log("Creating Work Order:", data);
    
    // Simulate API call
    setTimeout(() => {
      toast({
        title: "Work Order Created",
        description: `WO-2025-${String(data.wo_num).padStart(6, '0')} has been created successfully.`,
      });
    }, 1000);
  };

  return (
    <DashboardLayout currentPage="Create WO">
      <div className="space-y-6">
        <div>
          <h1 className="text-3xl font-bold tracking-tight">Create Work Order</h1>
          <p className="text-muted-foreground">
            Create a new work order to track pipe processing
          </p>
        </div>

        <CreateWOForm onSubmit={handleSubmit} />
      </div>
    </DashboardLayout>
  );
}