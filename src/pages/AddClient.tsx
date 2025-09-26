import { useState } from "react";
import { Header } from "@/components/layout/Header";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { ArrowLeft, UserPlus } from "lucide-react";
import { useNavigate } from "react-router-dom";
import { useAuth } from "@/contexts/AuthContext";
import { useSharePoint } from "@/contexts/SharePointContext";
import { useToast } from "@/hooks/use-toast";
import { safeLocalStorage } from "@/lib/safe-storage";

export default function AddClient() {
  const navigate = useNavigate();
  const { user } = useAuth();
  const { sharePointService, isConnected, refreshDataInBackground } = useSharePoint();
  const { toast } = useToast();

  const [formData, setFormData] = useState({
    client: "",
    payer: "",
  });
  const [isLoading, setIsLoading] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!user) {
      toast({ title: "Ошибка", description: "Пожалуйста, войдите в систему", variant: "destructive" });
      return;
    }
    if (!isConnected || !sharePointService) {
      toast({ title: "SharePoint не подключен", description: "Нажмите 'Load Data' в заголовке", variant: "destructive" });
      return;
    }
    if (!formData.client.trim()) {
      toast({ title: "Введите имя клиента", variant: "destructive" });
      return;
    }

    setIsLoading(true);
    try {
      const ok = await sharePointService.addClient(formData.client.trim(), formData.payer.trim());
      if (ok) {
        toast({ title: "✅ Клиент добавлен", description: `${formData.client} сохранён в листе client` });
        setFormData({ client: "", payer: "" });
        try {
          // Обновим кеш мгновенно
          safeLocalStorage.removeItem("sharepoint_last_refresh");
          if (refreshDataInBackground) await refreshDataInBackground(sharePointService);
        } catch (err) {
          console.warn("Failed to refresh cache after adding client:", err);
        }
      } else {
        toast({ title: "Не удалось сохранить", description: "Попробуйте снова через несколько секунд", variant: "destructive" });
      }
    } catch (err) {
      console.error(err);
      toast({ title: "Ошибка", description: "Не удалось добавить клиента", variant: "destructive" });
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50">
      <Header />
      <div className="container mx-auto px-6 py-8">
        <div className="mb-6">
          <Button variant="outline" onClick={() => navigate("/")} className="flex items-center space-x-2 border-2 hover:bg-gray-50">
            <ArrowLeft className="w-4 h-4" />
            <span>Back to Dashboard</span>
          </Button>
        </div>

        <Card className="max-w-2xl mx-auto border-2 shadow-lg">
          <CardHeader className="bg-indigo-50 border-b-2">
            <CardTitle className="text-2xl font-bold text-indigo-800 flex items-center gap-2">
              <UserPlus className="w-6 h-6" /> Add Client
            </CardTitle>
          </CardHeader>
          <CardContent className="p-6">
            <form onSubmit={handleSubmit} className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-2">
                  <Label htmlFor="client">Client Name</Label>
                  <Input id="client" className="h-11" value={formData.client} onChange={(e) => setFormData({ ...formData, client: e.target.value })} placeholder="Enter client name" required />
                </div>
                <div className="space-y-2">
                  <Label htmlFor="payer">Payer</Label>
                  <Input id="payer" className="h-11" value={formData.payer} onChange={(e) => setFormData({ ...formData, payer: e.target.value })} placeholder="Enter payer (optional)" />
                </div>
              </div>

              <div className="flex justify-end space-x-4 pt-6 border-t-2 border-gray-100">
                <Button type="button" variant="outline" onClick={() => navigate("/")} className="border-2 h-12 px-6">Cancel</Button>
                <Button type="submit" className="h-12 px-6 font-semibold" disabled={isLoading || !isConnected}>
                  {isLoading ? "Saving..." : "Save Client"}
                </Button>
              </div>
            </form>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
