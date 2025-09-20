import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import * as z from "zod";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import {
  Form,
  FormControl,
  FormField,
  FormItem,
  FormLabel,
  FormMessage,
} from "@/components/ui/form";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { BatchRegistryFormData } from "@/types";

const batchRegistrySchema = z.object({
  wo_id: z.string().min(1, "Work Order is required"),
  wo_num: z.number().min(1, "WO Number is required"),
  client: z.string().min(1, "Client is required"),
  qty: z.number().min(1, "Quantity must be greater than 0"),
  rack: z.string().min(1, "Rack is required"),
  note: z.string().optional(),
});

interface BatchRegistryFormProps {
  onSubmit: (data: BatchRegistryFormData) => void;
  isLoading?: boolean;
  workOrders?: Array<{ id: string; wo_num: number; client: string }>;
}

const racks = [
  "R-001", "R-002", "R-003", "R-004", "R-005",
  "R-006", "R-007", "R-008", "R-009", "R-010"
];

export function BatchRegistryForm({ 
  onSubmit, 
  isLoading = false, 
  workOrders = [] 
}: BatchRegistryFormProps) {
  const form = useForm<BatchRegistryFormData>({
    resolver: zodResolver(batchRegistrySchema),
    defaultValues: {
      wo_id: "",
      wo_num: 0,
      client: "",
      qty: 0,
      rack: "",
      note: "",
    },
  });

  const selectedWO = workOrders.find(wo => wo.id === form.watch("wo_id"));

  // Update client and wo_num when WO is selected
  const handleWOChange = (woId: string) => {
    const workOrder = workOrders.find(wo => wo.id === woId);
    if (workOrder) {
      form.setValue("wo_id", woId);
      form.setValue("wo_num", workOrder.wo_num);
      form.setValue("client", workOrder.client);
    }
  };

  return (
    <Card className="max-w-2xl mx-auto">
      <CardHeader>
        <CardTitle>Batch Registry</CardTitle>
      </CardHeader>
      <CardContent>
        <Form {...form}>
          <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-6">
            <FormField
              control={form.control}
              name="wo_id"
              render={({ field }) => (
                <FormItem>
                  <FormLabel>Work Order</FormLabel>
                  <Select onValueChange={handleWOChange} defaultValue={field.value}>
                    <FormControl>
                      <SelectTrigger>
                        <SelectValue placeholder="Select work order" />
                      </SelectTrigger>
                    </FormControl>
                    <SelectContent>
                      {workOrders.map((wo) => (
                        <SelectItem key={wo.id} value={wo.id}>
                          {wo.wo_num} • {wo.client}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                  <FormMessage />
                </FormItem>
              )}
            />

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <FormField
                control={form.control}
                name="client"
                render={({ field }) => (
                  <FormItem>
                    <FormLabel>Client</FormLabel>
                    <FormControl>
                      <Input {...field} disabled className="bg-muted" />
                    </FormControl>
                    <FormMessage />
                  </FormItem>
                )}
              />

              <FormField
                control={form.control}
                name="wo_num"
                render={({ field }) => (
                  <FormItem>
                    <FormLabel>WO Number</FormLabel>
                    <FormControl>
                      <Input 
                        {...field} 
                        disabled 
                        className="bg-muted"
                        value={field.value || ""}
                      />
                    </FormControl>
                    <FormMessage />
                  </FormItem>
                )}
              />

              <FormField
                control={form.control}
                name="qty"
                render={({ field }) => (
                  <FormItem>
                    <FormLabel>Quantity</FormLabel>
                    <FormControl>
                      <Input
                        type="number"
                        placeholder="Enter quantity"
                        {...field}
                        onChange={(e) => field.onChange(parseInt(e.target.value) || 0)}
                      />
                    </FormControl>
                    <FormMessage />
                  </FormItem>
                )}
              />

              <FormField
                control={form.control}
                name="rack"
                render={({ field }) => (
                  <FormItem>
                    <FormLabel>Rack</FormLabel>
                    <Select onValueChange={field.onChange} defaultValue={field.value}>
                      <FormControl>
                        <SelectTrigger>
                          <SelectValue placeholder="Select rack" />
                        </SelectTrigger>
                      </FormControl>
                      <SelectContent>
                        {racks.map((rack) => (
                          <SelectItem key={rack} value={rack}>
                            {rack}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                    <FormMessage />
                  </FormItem>
                )}
              />
            </div>

            <FormField
              control={form.control}
              name="note"
              render={({ field }) => (
                <FormItem>
                  <FormLabel>Note (Optional)</FormLabel>
                  <FormControl>
                    <Textarea 
                      placeholder="Enter any additional notes..."
                      className="resize-none"
                      {...field}
                    />
                  </FormControl>
                  <FormMessage />
                </FormItem>
              )}
            />

            <div className="flex justify-end space-x-4">
              <Button type="button" variant="outline">
                Cancel
              </Button>
              <Button type="submit" disabled={isLoading || !selectedWO}>
                {isLoading ? "Registering..." : "Register Batch"}
              </Button>
            </div>
          </form>
        </Form>
      </CardContent>
    </Card>
  );
}