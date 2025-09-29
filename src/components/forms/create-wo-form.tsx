import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import * as z from "zod";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Checkbox } from "@/components/ui/checkbox";
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

const createWOSchema = z.object({
  wo_num: z.number().min(1, "WO Number is required"),
  client: z.string().min(1, "Client is required"),
  pipe_type: z.string().min(1, "Pipe Type is required"),
  diameter: z.string().min(1, "Diameter is required"),
  coupling_replace: z.boolean(),
  transport: z.string().min(1, "Transport method is required"),
});

type CreateWOFormData = z.infer<typeof createWOSchema>;

interface CreateWOFormProps {
  onSubmit: (data: CreateWOFormData) => void;
  isLoading?: boolean;
}

const pipeTypes = [
  "Carbon Steel",
  "Stainless Steel", 
  "Alloy Steel",
  "Chrome",
  "Duplex"
];

const transportMethods = [
  "Truck",
  "Rail",
  "Ship",
  "Air",
  "Pipeline"
];

export function CreateWOForm({ onSubmit, isLoading = false }: CreateWOFormProps) {
  const form = useForm<CreateWOFormData>({
    resolver: zodResolver(createWOSchema),
    defaultValues: {
      wo_num: 0,
      client: "",
      pipe_type: "",
      diameter: "",
      coupling_replace: false,
      transport: "",
    },
  });

  return (
    <Card className="max-w-2xl mx-auto">
      <CardHeader>
        <CardTitle>Create Work Order</CardTitle>
      </CardHeader>
      <CardContent>
        <Form {...form}>
          <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <FormField
                control={form.control}
                name="wo_num"
                render={({ field }) => (
                  <FormItem>
                    <FormLabel>WO Number</FormLabel>
                    <FormControl>
                      <Input
                        type="number"
                        placeholder="Enter WO number"
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
                name="client"
                render={({ field }) => (
                  <FormItem>
                    <FormLabel>Client</FormLabel>
                    <FormControl>
                      <Input placeholder="Enter client name" {...field} />
                    </FormControl>
                    <FormMessage />
                  </FormItem>
                )}
              />

              <FormField
                control={form.control}
                name="pipe_type"
                render={({ field }) => (
                  <FormItem>
                    <FormLabel>Pipe Type</FormLabel>
                    <Select onValueChange={field.onChange} defaultValue={field.value}>
                      <FormControl>
                        <SelectTrigger>
                          <SelectValue placeholder="Select pipe type" />
                        </SelectTrigger>
                      </FormControl>
                      <SelectContent>
                        {pipeTypes.map((type) => (
                          <SelectItem key={type} value={type}>
                            {type}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                    <FormMessage />
                  </FormItem>
                )}
              />

              <FormField
                control={form.control}
                name="diameter"
                render={({ field }) => (
                  <FormItem>
                    <FormLabel>Diameter</FormLabel>
                    <FormControl>
                      <Input placeholder="e.g., 4.5 inches" {...field} />
                    </FormControl>
                    <FormMessage />
                  </FormItem>
                )}
              />

              <FormField
                control={form.control}
                name="transport"
                render={({ field }) => (
                  <FormItem>
                    <FormLabel>Transport</FormLabel>
                    <Select onValueChange={field.onChange} defaultValue={field.value}>
                      <FormControl>
                        <SelectTrigger>
                          <SelectValue placeholder="Select transport method" />
                        </SelectTrigger>
                      </FormControl>
                      <SelectContent>
                        {transportMethods.map((method) => (
                          <SelectItem key={method} value={method}>
                            {method}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                    <FormMessage />
                  </FormItem>
                )}
              />

              <FormField
                control={form.control}
                name="coupling_replace"
                render={({ field }) => (
                  <FormItem className="flex flex-row items-start space-x-3 space-y-0">
                    <FormControl>
                      <Checkbox
                        checked={field.value}
                        onCheckedChange={field.onChange}
                      />
                    </FormControl>
                    <div className="space-y-1 leading-none">
                      <FormLabel>Coupling Replace Required</FormLabel>
                    </div>
                  </FormItem>
                )}
              />
            </div>

            <div className="flex justify-end space-x-4">
              <Button type="button" variant="destructive">
                Cancel
              </Button>
              <Button type="submit" disabled={isLoading}>
                {isLoading ? "Creating..." : "Create Work Order"}
              </Button>
            </div>
          </form>
        </Form>
      </CardContent>
    </Card>
  );
}