import { useState } from "react";
import { Button } from "@/components/ui/button";
import { LogIn, LogOut, User, RefreshCw, TestTube, Download } from "lucide-react";
import { useAuth } from "@/contexts/AuthContext";
import { useSharePoint } from "@/contexts/SharePointContext";
import { LoginForm } from "@/components/auth/LoginForm";
import { authService } from "@/services/authService";
import { SharePointService } from "@/services/sharePointService";
import { DataStatusIndicator } from "@/components/ui/data-status-indicator";
import { safeLocalStorage } from '@/lib/safe-storage';

export const Header = () => {
  const { user, logout, isAuthenticated } = useAuth();
  const { 
    isConnected, 
    isConnecting, 
    connectToSharePoint, 
    refreshDataInBackground, 
    sharePointService,
    isDataLoading,
    cachedClients,
    cachedWorkOrders,
    resetExcelSession
  } = useSharePoint();
  const [showLoginForm, setShowLoginForm] = useState(false);
  
  console.log('Header render:', { user, isAuthenticated });

  const handleLogin = () => {
    setShowLoginForm(true);
  };

  const handleLogout = () => {
    logout();
  };

  const handleLoadData = async () => {
    if (!isConnected) {
      await connectToSharePoint();
    }
    if (sharePointService) {
      // Принудительное обновление при Load Data - очищаем время последнего обновления
      safeLocalStorage.removeItem("sharepoint_last_refresh");
      await refreshDataInBackground(sharePointService);
    }
  };

  const handleUpdateData = async () => {
    if (sharePointService) {
      // Принудительное обновление - очищаем время последнего обновления
      safeLocalStorage.removeItem("sharepoint_last_refresh");
      await refreshDataInBackground(sharePointService);
    }
  };

  const handleResetExcelSession = () => {
    try {
      resetExcelSession();
      console.log('Excel session reset requested by user');
    } catch (e) {
      console.error('Failed to reset Excel session:', e);
    }
  };

  const handleClearMSALCache = async () => {
    try {
      console.log('Clearing MSAL cache and forcing re-login...');
      await authService.initialize();
      const success = await authService.clearCacheAndRelogin();
      if (success) {
        console.log('Successfully cleared cache and re-logged in');
        alert('MSAL cache cleared and re-login successful!');
      } else {
        console.log('Re-login failed or was cancelled');
        alert('Re-login failed or was cancelled');
      }
    } catch (error) {
      console.error('Failed to clear MSAL cache:', error);
      alert('Failed to clear MSAL cache: ' + error);
    }
  };

  const handleTestSharePoint = async () => {
    try {
      console.log('Testing SharePoint access...');
      await authService.initialize();
      
      const graphClient = authService.getGraphClient();
      if (!graphClient) {
        alert('Not authenticated. Please login first.');
        return;
      }

      const authProvider = authService.getAuthenticationProvider();
      const spService = new SharePointService(authProvider);
      
      await spService.testSiteAccess();
      alert('SharePoint test completed! Check console for results.');
    } catch (error) {
      console.error('SharePoint test failed:', error);
      alert('SharePoint test failed: ' + error);
    }
  };

  return (
    <header className="border-b bg-white shadow-sm">
      <div className="container mx-auto px-4 py-3 flex items-center justify-between">
        {/* Logo */}
        <div className="flex items-center space-x-2">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center">
            <span className="text-white font-bold text-sm">FS</span>
          </div>
          <span className="text-xl font-semibold text-gray-900">FlowSync Forge</span>
        </div>
        
        {/* Authentication */}
        <div className="flex items-center space-x-3">
          {/* Data Control Buttons */}
          {!isConnected ? (
            <Button 
              onClick={handleLoadData}
              disabled={isConnecting}
              className="bg-green-600 hover:bg-green-700 text-white font-semibold px-4 py-2"
              size="sm"
            >
              <Download className="h-4 w-4 mr-2" />
              {isConnecting ? "Loading..." : "Load Data"}
            </Button>
          ) : (
            <div className="flex items-center space-x-2">
              <Button 
                onClick={handleUpdateData}
                disabled={isDataLoading}
                variant="outline"
                className="border-blue-500 text-blue-600 hover:bg-blue-50 font-semibold px-4 py-2"
                size="sm"
              >
                <RefreshCw className={`h-4 w-4 mr-2 ${isDataLoading ? 'animate-spin' : ''}`} />
                {isDataLoading ? "Updating..." : "Update Data"}
              </Button>
              <Button
                onClick={handleResetExcelSession}
                variant="outline"
                className="border-amber-500 text-amber-600 hover:bg-amber-50 font-semibold px-4 py-2"
                size="sm"
                title="Reset cached Excel workbook session"
              >
                <RefreshCw className="h-4 w-4 mr-2" />
                Reset Excel Session
              </Button>
            </div>
          )}
          
          {/* Data Status Indicator */}
          <DataStatusIndicator
            isConnected={isConnected}
            isLoading={isDataLoading}
            lastUpdate={safeLocalStorage.getItem("sharepoint_last_refresh") || undefined}
            dataCount={cachedClients.length + cachedWorkOrders.length}
          />
          
          {isAuthenticated ? (
            <>
              <div className="flex items-center space-x-2 text-sm text-gray-600">
                <User className="w-4 h-4" />
                <span>{user?.full_name} ({user?.role})</span>
              </div>
              <Button 
                variant="outline" 
                onClick={handleLogout}
                className="flex items-center space-x-2"
              >
                <LogOut className="w-4 h-4" />
                <span>Logout</span>
              </Button>
            </>
          ) : (
            <Button 
              variant="outline" 
              onClick={handleLogin}
              className="flex items-center space-x-2 bg-blue-600 text-white hover:bg-blue-700"
            >
              <LogIn className="w-4 h-4" />
              <span>Login</span>
            </Button>
          )}
        </div>
      </div>
      
      {showLoginForm && (
        <LoginForm onClose={() => setShowLoginForm(false)} />
      )}
    </header>
  );
};
