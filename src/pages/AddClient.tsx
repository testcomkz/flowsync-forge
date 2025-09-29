import { useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import { ArrowLeft, UserPlus } from "lucide-react";
import { Header } from "@/components/layout/Header";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { useAuth } from "@/contexts/AuthContext";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useToast } from "@/hooks/use-toast";
import { useSharePointInstantData } from "@/hooks/useInstantData";
import { safeLocalStorage } from "@/lib/safe-storage";
import { ConfirmDialog } from "@/components/ui/confirm-dialog";

export default function AddClient() {
  const navigate = useNavigate();
  const { user } = useAuth();
  const { isConnected, sharePointService, refreshDataInBackground } = useSharePoint();
  const { toast } = useToast();
  const { clientRecords } = useSharePointInstantData();

  const [formState, setFormState] = useState({
    name: "",
    payer: "",
  });
  const [isSaving, setIsSaving] = useState(false);
  const [isConfirmOpen, setIsConfirmOpen] = useState(false);

  const [confirmLines, setConfirmLines] = useState<string[]>([]);

  const doSave = async (name: string, payer: string) => {
    setIsSaving(true);
    try {
      const success = await sharePointService.addClientRecord({ name, payer });
      if (!success) {
        toast({
          title: "Не удалось сохранить",
          description: "Проверьте, что Excel файл свободен и попробуйте снова",
          variant: "destructive",
        });
        return;
      }

      const updatedRecords = await sharePointService.getClientRecordsFromExcel();
      const updatedNames = Array.from(new Set(updatedRecords.map(record => record.name).filter(Boolean))).sort((a, b) =>
        a.localeCompare(b)
      );
      const recordsPayload = JSON.stringify(updatedRecords);
      const namesPayload = JSON.stringify(updatedNames);
      const timestamp = new Date().toISOString();

      safeLocalStorage.setJSON("sharepoint_cached_client_records", updatedRecords);
      safeLocalStorage.setItem("sharepoint_cached_client_records_timestamp", timestamp);
      safeLocalStorage.dispatchStorageEvent("sharepoint_cached_client_records", recordsPayload);

      safeLocalStorage.setJSON("sharepoint_cached_clients", updatedNames);
      safeLocalStorage.setItem("sharepoint_cached_clients_timestamp", timestamp);
      safeLocalStorage.dispatchStorageEvent("sharepoint_cached_clients", namesPayload);

      toast({
        title: "Клиент добавлен",
        description: `${name} успешно сохранён в SharePoint`,
      });

      setFormState({ name: "", payer: "" });

      try {
        if (refreshDataInBackground) {
          safeLocalStorage.removeItem("sharepoint_last_refresh");
          await refreshDataInBackground(sharePointService);
        }
      } catch (error) {
        console.warn("Не удалось обновить кеш после добавления клиента:", error);
      }
    } catch (error) {
      console.error("Failed to add client:", error);
      toast({
        title: "Ошибка",
        description: "Не удалось добавить клиента. Попробуйте снова",
        variant: "destructive",
      });
    } finally {
      setIsSaving(false);
      setIsConfirmOpen(false);
    }
  };

  const sortedRecords = useMemo(() => {
    return [...clientRecords]
      .filter(record => record.name)
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [clientRecords]);

  const existingNames = useMemo(() => sortedRecords.map(record => record.name.toLowerCase()), [sortedRecords]);

  const handleInputChange = (field: "name" | "payer", value: string) => {
    setFormState(prev => ({ ...prev, [field]: value }));
  };

  const handleSubmit = async (event: React.FormEvent) => {
    event.preventDefault();

    if (!user) {
      toast({
        title: "Ошибка",
        description: "Пожалуйста, войдите в систему",
        variant: "destructive",
      });
      return;
    }

    if (!isConnected || !sharePointService) {
      toast({
        title: "SharePoint не подключен",
        description: "Подключитесь к SharePoint перед добавлением клиента",
        variant: "destructive",
      });
      return;
    }

    const name = formState.name.trim();
    const payer = formState.payer.trim();

    if (!name) {
      toast({
        title: "Ошибка валидации",
        description: "Введите имя клиента",
        variant: "destructive",
      });
      return;
    }

    if (!payer) {
      toast({
        title: "Ошибка валидации",
        description: "Введите Payer",
        variant: "destructive",
      });
      return;
    }

    if (existingNames.includes(name.toLowerCase())) {
      toast({
        title: "Клиент уже существует",
        description: "Этот клиент уже добавлен в список",
        variant: "destructive",
      });
      return;
    }

    setConfirmLines([
      `Client: ${name}`,
      `Payer: ${payer || '—'}`,
    ]);
    setIsConfirmOpen(true);
  };

  return (
    <div className="min-h-screen bg-slate-50">
      <Header />
      <div className="container mx-auto px-6 py-8">
        <div className="mb-6">
          <Button
            variant="ghost"
            onClick={() => navigate("/")}
            className="flex items-center gap-2 text-slate-600"
          >
            <ArrowLeft className="w-4 h-4" />
            <span>Back to Dashboard</span>
          </Button>
        </div>

        <Card className="max-w-4xl mx-auto border-2 border-blue-200 rounded-xl shadow-md">
          <CardHeader className="bg-blue-50 border-b flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <div>
              <CardTitle className="text-2xl font-bold text-blue-800">Add Client</CardTitle>
              <p className="text-sm text-blue-600">Создайте нового клиента и свяжите с ним Payer</p>
            </div>
            <div className="hidden md:flex items-center justify-center w-14 h-14 rounded-full bg-white shadow-md border border-blue-100">
              <UserPlus className="w-7 h-7 text-blue-600" />
            </div>
          </CardHeader>
          <CardContent className="p-6 space-y-10">
            <ConfirmDialog
              open={isConfirmOpen}
              title="Add Client?"
              description="Please confirm adding the following client to SharePoint"
              lines={confirmLines}
              confirmText="Add"
              cancelText="Cancel"
              onConfirm={() => doSave(formState.name.trim(), formState.payer.trim())}
              onCancel={() => setIsConfirmOpen(false)}
              loading={isSaving}
            />
            <form onSubmit={handleSubmit} className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-2">
                  <Label htmlFor="client_name" className="text-sm font-semibold text-gray-700">Client Name</Label>
                  <Input
                    id="client_name"
                    value={formState.name}
                    onChange={(event) => handleInputChange("name", event.target.value)}
                    placeholder="Введите имя клиента"
                    className="border-2 focus:border-blue-500 h-11"
                    required
                  />
                </div>
                <div className="space-y-2">
                  <Label htmlFor="payer" className="text-sm font-semibold text-gray-700">Payer</Label>
                  <Input
                    id="payer"
                    value={formState.payer}
                    onChange={(event) => handleInputChange("payer", event.target.value)}
                    placeholder="Введите Payer"
                    className="border-2 focus:border-blue-500 h-11"
                    required
                  />
                </div>
              </div>
              <div className="flex justify-end">
                <Button type="submit" className="h-11 px-6" disabled={isSaving || !isConnected}>
                  {isSaving ? "Saving..." : "Save Client"}
                </Button>
              </div>
            </form>

            <div className="space-y-4">
              <h2 className="text-lg font-semibold text-gray-800">Existing Clients</h2>
              <div className="border rounded-lg divide-y bg-white">
                {sortedRecords.length === 0 ? (
                  <div className="p-4 text-sm text-gray-500">Клиенты ещё не добавлены</div>
                ) : (
                  sortedRecords.map(record => (
                    <div key={`${record.name}-${record.payer}`} className="grid grid-cols-1 md:grid-cols-[2fr_1fr] gap-4 p-4">
                      <div>
                        <p className="text-sm font-semibold text-gray-800">{record.name}</p>
                        <p className="text-xs text-gray-500">Client</p>
                      </div>
                      <div>
                        <p className="text-sm font-semibold text-gray-800">{record.payer || "—"}</p>
                        <p className="text-xs text-gray-500">Payer</p>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
