import { Configuration } from "@azure/msal-browser";
import { safeLocalStorage } from '@/lib/safe-storage';

export const msalConfig: Configuration = {
  auth: {
    clientId: import.meta.env.VITE_MSAL_CLIENT_ID,
    authority: import.meta.env.VITE_MSAL_AUTHORITY,
    redirectUri: import.meta.env.VITE_MSAL_REDIRECT_URI,
  },
  cache: {
    cacheLocation: "localStorage", // Изменено с sessionStorage на localStorage для постоянного хранения
    storeAuthStateInCookie: true, // Включено для дополнительной персистентности
  },
};

export const loginRequest = {
  scopes: ["User.Read", "Sites.ReadWrite.All", "Files.ReadWrite.All", "offline_access"],
  prompt: "none" // Изменено с "select_account" на "none" для автоматического входа
};

export const graphConfig = {
  graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
  sharePointSiteUrl: "https://kzprimeestate.sharepoint.com/sites/pipe-inspection",
};

// Utility function to clear MSAL cache and force re-login
export const clearMSALCache = () => {
  console.log('Clearing MSAL cache...');
  
  // Clear localStorage (where MSAL cache is now stored)
  const msalKeys = safeLocalStorage
    .keys()
    .filter(key =>
      key.startsWith('msal.') ||
      key.includes('msal') ||
      key.startsWith('sharepoint_')
    );

  msalKeys.forEach(key => {
    safeLocalStorage.removeItem(key);
    console.log(`Removed: ${key}`);
  });

  // Clear sessionStorage as well
  if (typeof window !== "undefined" && window.sessionStorage) {
    window.sessionStorage.clear();
  }

  // Clear any cookies related to authentication
  if (typeof document !== "undefined") {
    document.cookie.split(";").forEach((c) => {
      const eqPos = c.indexOf("=");
      const name = eqPos > -1 ? c.substr(0, eqPos) : c;
      document.cookie = name + "=;expires=Thu, 01 Jan 1970 00:00:00 GMT;path=/";
    });
  }
  
  console.log('MSAL cache cleared. Please refresh the page to force re-login.');
};
