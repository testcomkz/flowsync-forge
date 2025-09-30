import { authService } from '@/services/authService';

export async function getSharePointFileId(): Promise<string | null> {
  try {
    await authService.initialize();
    
    const siteId = import.meta.env.VITE_SP_SITE_ID;
    const filePath = import.meta.env.VITE_SP_FILE_PATH;
    
    const token = await authService.getAccessToken();
    if (!token) throw new Error('No access token');

    // Парсим site ID (берем последнюю часть)
    const siteIdParts = siteId.split(',');
    const actualSiteId = siteIdParts.length === 3 
      ? `${siteIdParts[0]},${siteIdParts[1]},${siteIdParts[2]}`
      : siteId;

    const response = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${actualSiteId}/drive/root:${filePath}`,
      {
        headers: {
          'Authorization': `Bearer ${token}`
        }
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to get file: ${response.statusText}`);
    }

    const data = await response.json();
    console.log('📄 File ID:', data.id);
    console.log('📄 Full file info:', data);
    
    return data.id;
  } catch (error) {
    console.error('❌ Error getting file ID:', error);
    return null;
  }
}
