import { PublicClientApplication, Configuration } from "@azure/msal-browser";
import { Client, AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import { msalConfig, loginRequest, clearMSALCache } from "@/config/authConfig";

// Graph API scopes - full permissions for SharePoint access
const graphScopes = loginRequest.scopes;

class AuthService {
  private msalInstance: PublicClientApplication;
  private graphClient: Client | null = null;

  constructor() {
    this.msalInstance = new PublicClientApplication(msalConfig);
  }

  // Инициализация MSAL
  async initialize(): Promise<void> {
    await this.msalInstance.initialize();
  }

  // Вход в систему
  async login(): Promise<boolean> {
    try {
      // Сначала пробуем тихий вход (silent login) - ВСЕГДА
      const accounts = this.msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        try {
          const silentRequest = {
            scopes: graphScopes,
            account: accounts[0],
          };
          const silentResult = await this.msalInstance.acquireTokenSilent(silentRequest);
          if (silentResult) {
            this.msalInstance.setActiveAccount(accounts[0]);
            await this.initializeGraphClient();
            console.log('✅ Silent login successful');
            return true;
          }
        } catch (silentError) {
          console.log('Silent login failed:', silentError);
          // НЕ показываем popup сразу - возможно это временная проблема
          return false;
        }
      }

      // Popup только если явно запрошен пользователем
      console.log('No valid tokens found, popup required');
      return false;
    } catch (error) {
      console.error("Login failed:", error);
      return false;
    }
  }

  // Отдельный метод для принудительного popup входа
  async forcePopupLogin(): Promise<boolean> {
    try {
      const loginResponse = await this.msalInstance.loginPopup({
        scopes: graphScopes,
        prompt: "select_account",
      });

      if (loginResponse) {
        this.msalInstance.setActiveAccount(loginResponse.account);
        await this.initializeGraphClient();
        console.log('✅ Popup login successful');
        return true;
      }
      return false;
    } catch (error) {
      console.error("Popup login failed:", error);
      return false;
    }
  }

  // Выход из системы
  async logout(): Promise<void> {
    await this.msalInstance.logoutPopup();
    this.graphClient = null;
  }

  // Очистка кеша и принудительный повторный вход
  async clearCacheAndRelogin(): Promise<boolean> {
    try {
      // Clear MSAL cache
      clearMSALCache();
      
      // Reinitialize MSAL instance
      this.msalInstance = new PublicClientApplication(msalConfig);
      await this.msalInstance.initialize();
      
      // Force login with account selection
      const silentSuccess = await this.login();
      if (silentSuccess) {
        return true;
      }

      // If there are no active accounts after cache clear, force popup login
      return await this.forcePopupLogin();
    } catch (error) {
      console.error("Failed to clear cache and relogin:", error);
      return false;
    }
  }

  // Проверка аутентификации
  isAuthenticated(): boolean {
    const accounts = this.msalInstance.getAllAccounts();
    return accounts.length > 0;
  }

  // Получить Graph Client
  getGraphClient(): Client | null {
    return this.graphClient;
  }

  // Получить AuthenticationProvider для Microsoft Graph
  getAuthenticationProvider(): AuthenticationProvider {
    const provider: AuthenticationProvider = {
      getAccessToken: async () => {
        try {
          const accounts = this.msalInstance.getAllAccounts();
          const account = this.msalInstance.getActiveAccount() || accounts[0];
          
          if (!account) {
            console.warn('⚠️ No MSAL account found for SharePoint access');
            throw new Error("No active MSAL account. Please connect to SharePoint first.");
          }

          console.log(`🔑 Getting access token for account: ${account.username}`);
          
          try {
            const response = await this.msalInstance.acquireTokenSilent({
              scopes: graphScopes,
              account,
            });
            
            if (!response.accessToken) {
              throw new Error('Access token is empty from MSAL');
            }
            
            console.log(`✅ Successfully got access token for SharePoint`);
            return response.accessToken;
          } catch (silentError: any) {
            console.error('❌ Silent token acquisition failed:', silentError);
            
            // Если silent не работает, пробуем интерактивный способ
            if (silentError.errorCode === 'consent_required' || 
                silentError.errorCode === 'interaction_required' ||
                silentError.errorCode === 'login_required') {
              
              console.log('🔄 Attempting interactive token acquisition...');
              try {
                const interactiveResponse = await this.msalInstance.acquireTokenPopup({
                  scopes: graphScopes,
                  account,
                });
                
                if (!interactiveResponse.accessToken) {
                  throw new Error('Interactive token acquisition returned empty token');
                }
                
                console.log(`✅ Successfully got access token via popup`);
                return interactiveResponse.accessToken;
              } catch (interactiveError) {
                console.error('❌ Interactive token acquisition failed:', interactiveError);
                throw new Error(`SharePoint authentication failed: ${interactiveError}`);
              }
            } else {
              throw new Error(`Token acquisition failed: ${silentError.message}`);
            }
          }
        } catch (error) {
          console.error('❌ getAccessToken failed:', error);
          throw error;
        }
      },
    };
    return provider;
  }

  // Инициализация Graph Client
  private async initializeGraphClient(): Promise<void> {
    const authProvider = this.getAuthenticationProvider();
    this.graphClient = Client.initWithMiddleware({ authProvider });
  }

  // Получить токен доступа
  async getAccessToken(): Promise<string | null> {
    try {
      const accounts = this.msalInstance.getAllAccounts();
      if (accounts.length === 0) return null;

      const response = await this.msalInstance.acquireTokenSilent({
        scopes: graphScopes,
        account: accounts[0],
      });

      return response.accessToken;
    } catch (error) {
      console.error("Failed to get access token:", error);
      return null;
    }
  }
}

export const authService = new AuthService();
