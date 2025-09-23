import { Controller, useForm } from "react-hook-form";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { useToast } from "@/hooks/use-toast";
import { DateInputField } from "@/components/ui/date-input";

interface InspectionFormData {
  client: string;
  wo_no: number;
  batch: string;
  diameter: string;
  qty: number;
  pipe_from: number;
  pipe_to: number;
  class_1: number;
  class_2: number;
  class_3: number;
  repair: number;
  scrap: number;
  start_date: string;
  end_date: string;
  rattling_qty: number;
  external_qty: number;
  hydro_qty: number;
  mpi_qty: number;
  drift_qty: number;
  emi_qty: number;
  marking_qty: number;
  act_no_oper: string;
  act_date: string;
}

export function InspectionForm() {
  const { register, handleSubmit, reset, control, formState: { errors } } = useForm<InspectionFormData>({
    defaultValues: {
      start_date: "",
      end_date: "",
      act_date: "",
    },
  });
  const { toast } = useToast();

  const onSubmit = (data: InspectionFormData) => {
    console.log("Inspection data:", data);
    
    toast({
      title: "Record Saved",
      description: "Inspection record has been saved successfully.",
    });
    
    reset();
  };

  return (
    <Card>
      <CardHeader>
        <CardTitle>Inspection Record Entry</CardTitle>
      </CardHeader>
      <CardContent>
        <form onSubmit={handleSubmit(onSubmit)} className="space-y-6">
          {/* Basic Information */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <Label htmlFor="client">Client</Label>
              <Input
                id="client"
                {...register("client", { required: "Client is required" })}
                placeholder="Enter client name"
              />
              {errors.client && (
                <p className="text-sm text-destructive mt-1">{errors.client.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="wo_no">WO No</Label>
              <Input
                id="wo_no"
                type="number"
                {...register("wo_no", { required: "WO No is required", valueAsNumber: true })}
                placeholder="Enter work order number"
              />
              {errors.wo_no && (
                <p className="text-sm text-destructive mt-1">{errors.wo_no.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="batch">Batch</Label>
              <Input
                id="batch"
                {...register("batch", { required: "Batch is required" })}
                placeholder="Enter batch"
              />
              {errors.batch && (
                <p className="text-sm text-destructive mt-1">{errors.batch.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="diameter">Diameter</Label>
              <Input
                id="diameter"
                {...register("diameter", { required: "Diameter is required" })}
                placeholder="Enter diameter"
              />
              {errors.diameter && (
                <p className="text-sm text-destructive mt-1">{errors.diameter.message}</p>
              )}
            </div>
          </div>

          {/* Quantities */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div>
              <Label htmlFor="qty">Qty</Label>
              <Input
                id="qty"
                type="number"
                {...register("qty", { required: "Qty is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.qty && (
                <p className="text-sm text-destructive mt-1">{errors.qty.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="pipe_from">Pipe From</Label>
              <Input
                id="pipe_from"
                type="number"
                {...register("pipe_from", { required: "Pipe From is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.pipe_from && (
                <p className="text-sm text-destructive mt-1">{errors.pipe_from.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="pipe_to">Pipe To</Label>
              <Input
                id="pipe_to"
                type="number"
                {...register("pipe_to", { required: "Pipe To is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.pipe_to && (
                <p className="text-sm text-destructive mt-1">{errors.pipe_to.message}</p>
              )}
            </div>
          </div>

          {/* Classification */}
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
            <div>
              <Label htmlFor="class_1">Class 1</Label>
              <Input
                id="class_1"
                type="number"
                {...register("class_1", { required: "Class 1 is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.class_1 && (
                <p className="text-sm text-destructive mt-1">{errors.class_1.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="class_2">Class 2</Label>
              <Input
                id="class_2"
                type="number"
                {...register("class_2", { required: "Class 2 is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.class_2 && (
                <p className="text-sm text-destructive mt-1">{errors.class_2.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="class_3">Class 3</Label>
              <Input
                id="class_3"
                type="number"
                {...register("class_3", { required: "Class 3 is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.class_3 && (
                <p className="text-sm text-destructive mt-1">{errors.class_3.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="repair">Repair</Label>
              <Input
                id="repair"
                type="number"
                {...register("repair", { required: "Repair is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.repair && (
                <p className="text-sm text-destructive mt-1">{errors.repair.message}</p>
              )}
            </div>
          </div>

          {/* Additional Fields */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <Label htmlFor="scrap">Scrap</Label>
              <Input
                id="scrap"
                type="number"
                {...register("scrap", { required: "Scrap is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.scrap && (
                <p className="text-sm text-destructive mt-1">{errors.scrap.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="rattling_qty">Rattling Qty</Label>
              <Input
                id="rattling_qty"
                type="number"
                {...register("rattling_qty", { required: "Rattling Qty is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.rattling_qty && (
                <p className="text-sm text-destructive mt-1">{errors.rattling_qty.message}</p>
              )}
            </div>
          </div>

          {/* Inspection Types */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div>
              <Label htmlFor="external_qty">External Qty</Label>
              <Input
                id="external_qty"
                type="number"
                {...register("external_qty", { required: "External Qty is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.external_qty && (
                <p className="text-sm text-destructive mt-1">{errors.external_qty.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="hydro_qty">Hydro Qty</Label>
              <Input
                id="hydro_qty"
                type="number"
                {...register("hydro_qty", { required: "Hydro Qty is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.hydro_qty && (
                <p className="text-sm text-destructive mt-1">{errors.hydro_qty.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="mpi_qty">MPI Qty</Label>
              <Input
                id="mpi_qty"
                type="number"
                {...register("mpi_qty", { required: "MPI Qty is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.mpi_qty && (
                <p className="text-sm text-destructive mt-1">{errors.mpi_qty.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="drift_qty">Drift Qty</Label>
              <Input
                id="drift_qty"
                type="number"
                {...register("drift_qty", { required: "Drift Qty is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.drift_qty && (
                <p className="text-sm text-destructive mt-1">{errors.drift_qty.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="emi_qty">EMI Qty</Label>
              <Input
                id="emi_qty"
                type="number"
                {...register("emi_qty", { required: "EMI Qty is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.emi_qty && (
                <p className="text-sm text-destructive mt-1">{errors.emi_qty.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="marking_qty">Marking Qty</Label>
              <Input
                id="marking_qty"
                type="number"
                {...register("marking_qty", { required: "Marking Qty is required", valueAsNumber: true })}
                placeholder="0"
              />
              {errors.marking_qty && (
                <p className="text-sm text-destructive mt-1">{errors.marking_qty.message}</p>
              )}
            </div>
          </div>

          {/* Dates and Additional Info */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <Label htmlFor="start_date">Start Date</Label>
              <Controller
                control={control}
                name="start_date"
                rules={{ required: "Start Date is required" }}
                render={({ field }) => (
                  <DateInputField
                    id="start_date"
                    value={field.value}
                    onChange={field.onChange}
                    placeholder="dd/mm/yyyy"
                  />
                )}
              />
              {errors.start_date && (
                <p className="text-sm text-destructive mt-1">{errors.start_date.message}</p>
              )}
            </div>

            <div>
              <Label htmlFor="end_date">End Date</Label>
              <Controller
                control={control}
                name="end_date"
                rules={{ required: "End Date is required" }}
                render={({ field }) => (
                  <DateInputField
                    id="end_date"
                    value={field.value}
                    onChange={field.onChange}
                    placeholder="dd/mm/yyyy"
                  />
                )}
              />
              {errors.end_date && (
                <p className="text-sm text-destructive mt-1">{errors.end_date.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="act_no_oper">Act No Oper</Label>
              <Input
                id="act_no_oper"
                {...register("act_no_oper", { required: "Act No Oper is required" })}
                placeholder="Enter act no oper"
              />
              {errors.act_no_oper && (
                <p className="text-sm text-destructive mt-1">{errors.act_no_oper.message}</p>
              )}
            </div>
            
            <div>
              <Label htmlFor="act_date">Act Date</Label>
              <Controller
                control={control}
                name="act_date"
                rules={{ required: "Act Date is required" }}
                render={({ field }) => (
                  <DateInputField
                    id="act_date"
                    value={field.value}
                    onChange={field.onChange}
                    placeholder="dd/mm/yyyy"
                  />
                )}
              />
              {errors.act_date && (
                <p className="text-sm text-destructive mt-1">{errors.act_date.message}</p>
              )}
            </div>
          </div>

          <div className="flex gap-4 pt-4">
            <Button type="submit" className="flex-1">
              Save Record
            </Button>
            <Button type="button" variant="outline" onClick={() => reset()}>
              Clear Form
            </Button>
          </div>
        </form>
      </CardContent>
    </Card>
  );
}