import { useState, useEffect } from "react";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { FileText, Database, Edit3, Lock, FileSpreadsheet, Wifi, WifiOff, ClipboardCheck } from "lucide-react";
import { useNavigate } from "react-router-dom";
import { useAuth } from "@/contexts/AuthContext";
import { useSharePoint } from "@/contexts/SharePointContext";
import { Header } from "@/components/layout/Header";
import { useToast } from "@/hooks/use-toast";

export const MainDashboard = () => {
  const navigate = useNavigate();
  const { user, isAuthenticated, isLoading } = useAuth();
  const { toast } = useToast();
  const { isConnected, isConnecting, connect, disconnect, error, ensureLatestData } = useSharePoint();

  // Показываем загрузку пока проверяем аутентификацию
  if (isLoading) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
          <p className="text-gray-600">Загрузка...</p>
        </div>
      </div>
    );
  }

  const handleCardClick = async (path: string) => {
    if (!isAuthenticated) {
      alert("Пожалуйста, войдите в систему для доступа к этой функции");
      return;
    }
    if (!isConnected) {
      alert("Сначала подключитесь к SharePoint");
      return;
    }
    try {
      await ensureLatestData();
    } catch (err) {
      console.warn('Failed to refresh data before navigation:', err);
    }
    navigate(path);
  };

  const handleSharePointConnect = async () => {
    const success = await connect();
    if (success) {
      toast({
        title: "Подключение успешно",
        description: "SharePoint подключен и сохранен",
        });
    } else {
      toast({
        title: "Connection Failed",
        description: "Failed to connect to SharePoint",
        variant: "destructive",
      });
    }
  };

  const handleSharePointDisconnect = () => {
    disconnect();
    toast({
      title: "Disconnected",
      description: "SharePoint disconnected",
    });
  };

  // Удаляем старый код с ошибками
  const oldHandleConnect = () => {
    // Старый код удален
  };

  const workCards = [
    {
      title: "Add Work Order",
      description: "Create new work orders with client and project details",
      icon: FileText,
      action: () => handleCardClick("/wo-form"),
      color: "bg-blue-50 hover:bg-blue-100 border-blue-300"
    },
    {
      title: "Tubing Registry",
      description: "Register and track tubing batches and inspections",
      icon: Database,
      action: () => handleCardClick("/tubing-form"),
      color: "bg-green-50 hover:bg-green-100 border-green-300"
    },
    {
      title: "Inspection Data",
      description: "Complete inspection details for arrived batches",
      icon: ClipboardCheck,
      action: () => handleCardClick("/inspection-data"),
      color: "bg-emerald-50 hover:bg-emerald-100 border-emerald-300"
    },
    {
      title: "Edit Records",
      description: "Modify existing work orders and tubing records",
      icon: Edit3,
      action: () => handleCardClick("/edit"),
      color: "bg-orange-50 hover:bg-orange-100 border-orange-300"
    }
  ];

  return (
    <div className="min-h-screen bg-gray-50">
      <Header />
      <div className="container mx-auto px-6 py-8">
      <div className="mb-8 text-center">
        <h1 className="text-3xl font-bold text-gray-900 mb-2">Work Registry Dashboard</h1>
        <p className="text-gray-600">Manage work orders and tubing inspections</p>
        {!isAuthenticated && (
          <div className="mt-4 p-4 bg-yellow-50 border border-yellow-200 rounded-lg">
            <div className="flex items-center justify-center space-x-2 text-yellow-800">
              <Lock className="w-5 h-5" />
              <span>Войдите в систему для доступа к функциям</span>
            </div>
          </div>
        )}
        {isAuthenticated && !isConnected && (
          <div className="mt-4 p-4 bg-purple-50 border border-purple-200 rounded-lg">
            <div className="flex items-center justify-center space-x-2 text-purple-800">
              <WifiOff className="w-5 h-5" />
              <span>Подключитесь к SharePoint для доступа к рабочим функциям</span>
            </div>
          </div>
        )}
      </div>
      
      <div className={`grid gap-6 max-w-6xl mx-auto ${
        !isAuthenticated 
          ? 'grid-cols-1 max-w-md' 
          : isAuthenticated && !isConnected 
            ? 'grid-cols-1 max-w-md'
            : 'grid-cols-1 md:grid-cols-2 lg:grid-cols-3'
      }`}>
        {/* SharePoint Connection Card - Показывается только после входа */}
        {isAuthenticated && (
          <Card className={`cursor-pointer transition-all duration-200 border-2 shadow-lg hover:shadow-xl ${
            isConnected 
              ? 'bg-green-50 hover:bg-green-100 border-green-300' 
              : 'bg-purple-50 hover:bg-purple-100 border-purple-300'
          } ${
            !isAuthenticated ? 'opacity-50 cursor-not-allowed' : ''
          }`}>
          <CardHeader className="text-center pb-4">
            <div className="mx-auto w-14 h-14 rounded-full bg-white flex items-center justify-center mb-4 shadow-md border-2 border-gray-100">
              {isConnecting ? (
                <FileSpreadsheet className="w-7 h-7 text-gray-400 animate-spin" />
              ) : isConnected ? (
                <Wifi className="w-7 h-7 text-green-600" />
              ) : (
                <WifiOff className="w-7 h-7 text-purple-600" />
              )}
            </div>
            <CardTitle className="text-xl font-bold">
              {isConnected ? "SharePoint Connected" : "Connect to SharePoint"}
            </CardTitle>
            <CardDescription className="text-base text-gray-700 font-medium">
              {isConnected 
                ? "Access pipe inspection data from SharePoint Excel" 
                : "Connect to SharePoint to access work order data"}
            </CardDescription>
          </CardHeader>
          <CardContent className="pt-0">
            <Button 
              className="w-full h-12 text-base font-semibold border-2" 
              variant={isConnected ? "outline" : "default"}
              disabled={isConnecting || !isAuthenticated}
              onClick={isConnected ? () => navigate("/sharepoint-viewer") : handleSharePointConnect}
            >
              {isConnecting ? "Connecting..." : 
               isConnected ? "Open SharePoint Viewer" : "Connect to SharePoint"}
            </Button>
          </CardContent>
        </Card>
        )}

        {/* Work Cards - Only show when SharePoint is connected */}
        {isConnected && workCards.map((card, index) => {
          const IconComponent = card.icon;
          return (
            <Card
              key={index}
              className={`cursor-pointer transition-all duration-200 border-2 shadow-lg hover:shadow-xl ${card.color} ${!isAuthenticated ? 'opacity-60' : ''}`}
              onClick={() => {
                void card.action();
              }}
            >
              <CardHeader className="text-center pb-4">
                <div className="mx-auto w-14 h-14 rounded-full bg-white flex items-center justify-center mb-4 shadow-md border-2 border-gray-100">
                  {!isAuthenticated ? (
                    <Lock className="w-7 h-7 text-gray-400" />
                  ) : (
                    <IconComponent className="w-7 h-7 text-gray-700" />
                  )}
                </div>
                <CardTitle className="text-xl font-bold">{card.title}</CardTitle>
                <CardDescription className="text-base text-gray-700 font-medium">
                  {card.description}
                </CardDescription>
              </CardHeader>
              <CardContent className="pt-0">
                <Button
                  className="w-full h-12 text-base font-semibold border-2"
                  variant={isAuthenticated ? "default" : "outline"}
                  disabled={!isAuthenticated}
                  onClick={async (e) => {
                    e.stopPropagation();
                    await card.action();
                  }}
                >
                  {isAuthenticated ? `Open ${card.title}` : "Login Required"}
                </Button>
              </CardContent>
            </Card>
          );
        })}
      </div>
      </div>
    </div>
  );
};
