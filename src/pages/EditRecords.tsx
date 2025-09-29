import { useNavigate } from "react-router-dom";
import { ArrowLeft, ClipboardEdit, Layers, Wrench, Users } from "lucide-react";

import { Header } from "@/components/layout/Header";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";

export default function EditRecords() {
  const navigate = useNavigate();

  return (
    <div className="min-h-screen bg-slate-50">
      <Header />
      <main className="container mx-auto px-4 py-6">
        <div className="mb-6 flex items-center justify-between">
          <Button variant="ghost" onClick={() => navigate("/")} className="flex items-center gap-2 text-slate-600">
            <ArrowLeft className="h-4 w-4" />
            Back
          </Button>
          <div className="flex items-center gap-2 text-sm text-muted-foreground">
            <ClipboardEdit className="h-4 w-4 text-blue-500" />
            Edit Records
          </div>
        </div>

        <div className="grid gap-4 md:grid-cols-3">
          <Card className="border-2 border-blue-200 bg-white hover:shadow-md transition rounded-xl">
            <CardContent className="p-6 flex flex-col items-center gap-4">
              <div className="w-12 h-12 rounded-full bg-white ring-1 ring-blue-200 shadow-sm flex items-center justify-center">
                <Layers className="h-6 w-6 text-blue-600" />
              </div>
              <div className="text-center">
                <p className="text-lg font-semibold text-slate-900">Batch Edit</p>
                <p className="text-sm text-slate-600">Batch editing workflow</p>
              </div>
              <Button onClick={() => navigate('/batch-edit')} className="mt-1 bg-blue-600 hover:bg-blue-700 text-white w-full">Open Batch Edit</Button>
            </CardContent>
          </Card>

          <Card className="border-2 border-blue-200 bg-white hover:shadow-md transition rounded-xl">
            <CardContent className="p-6 flex flex-col items-center gap-4">
              <div className="w-12 h-12 rounded-full bg-white ring-1 ring-blue-200 shadow-sm flex items-center justify-center">
                <Wrench className="h-6 w-6 text-blue-600" />
              </div>
              <div className="text-center">
                <p className="text-lg font-semibold text-slate-900">Edit Work Orders</p>
                <p className="text-sm text-slate-600">Edit WO fields</p>
              </div>
              <Button onClick={() => navigate('/workorder-edit-select')} className="mt-1 bg-blue-600 hover:bg-blue-700 text-white w-full">Open Edit Work Orders</Button>
            </CardContent>
          </Card>

          <Card className="border-2 border-blue-200 bg-white hover:shadow-md transition rounded-xl">
            <CardContent className="p-6 flex flex-col items-center gap-4">
              <div className="w-12 h-12 rounded-full bg-white ring-1 ring-blue-200 shadow-sm flex items-center justify-center">
                <Users className="h-6 w-6 text-blue-600" />
              </div>
              <div className="text-center">
                <p className="text-lg font-semibold text-slate-900">Edit Clients</p>
                <p className="text-sm text-slate-600">Manage client names & payer</p>
              </div>
              <Button onClick={() => navigate('/edit-clients')} className="mt-1 bg-blue-600 hover:bg-blue-700 text-white w-full">Open Edit Clients</Button>
            </CardContent>
          </Card>
        </div>
      </main>
    </div>
  );
}
