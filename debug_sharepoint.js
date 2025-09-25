// Выполните этот код в Console браузера (F12 → Console) на странице вашего приложения
// где пользователь уже авторизован в SharePoint

async function getSharePointIds() {
  try {
    // Получаем токен из существующего authService
    const token = await authService.getAccessToken();
    console.log('🔑 Access Token:', token ? 'Present' : 'Missing');
    
    if (!token) {
      console.error('❌ No access token! Please connect to SharePoint first.');
      return;
    }

    const siteId = 'kzprimeestate.sharepoint.com,9f482633-8093-471a-a0e4-c6353a265373,df166f91-e314-4b9b-87a7-cbd2e26f2c1c';
    const headers = { 'Authorization': `Bearer ${token}` };
    
    console.log('📍 Step 1: Getting drives...');
    
    // 1. Получаем список drives
    const drivesResponse = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=id,name,driveType,webUrl`,
      { headers }
    );
    const drivesData = await drivesResponse.json();
    console.log('📁 Available drives:', drivesData.value);
    
    // Ищем Documents drive
    const documentsLibrary = drivesData.value.find(d => d.name === 'Documents' || d.name === 'Документы');
    if (!documentsLibrary) {
      console.error('❌ Documents library not found!');
      return;
    }
    
    const driveId = documentsLibrary.id;
    console.log('✅ DRIVE_ID:', driveId);
    
    console.log('📍 Step 2: Finding Excel file...');
    
    // 2. Поиск Excel файла
    const filePath = '/UPLOADS/pipe-inspection.xlsm';
    const fileResponse = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:${filePath}?$select=id,name,webUrl`,
      { headers }
    );
    
    if (!fileResponse.ok) {
      console.error('❌ File not found at path:', filePath);
      
      // Попробуем поиск по имени
      console.log('🔍 Searching by filename...');
      const searchResponse = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root/search(q='pipe_inspection.xlsm')?$select=id,name,webUrl,parentReference`,
        { headers }
      );
      const searchData = await searchResponse.json();
      console.log('🔍 Search results:', searchData.value);
      
      if (searchData.value.length === 0) {
        console.error('❌ File pipe_inspection.xlsm not found anywhere!');
        return;
      }
      
      const foundFile = searchData.value[0];
      const itemId = foundFile.id;
      console.log('✅ ITEM_ID (from search):', itemId);
      console.log('📂 File location:', foundFile.parentReference?.path);
      
    } else {
      const fileData = await fileResponse.json();
      const itemId = fileData.id;
      console.log('✅ ITEM_ID:', itemId);
    }
    
    console.log('📍 Step 3: Testing Excel access...');
    
    // 3. Проверяем доступ к worksheets
    const worksheetsResponse = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${itemId}/workbook/worksheets?$select=name`,
      { headers }
    );
    
    if (worksheetsResponse.ok) {
      const worksheetsData = await worksheetsResponse.json();
      console.log('📊 Available worksheets:', worksheetsData.value.map(w => w.name));
    } else {
      console.error('❌ Cannot access Excel worksheets:', await worksheetsResponse.text());
    }
    
    console.log('\n🎯 FINAL RESULTS:');
    console.log(`VITE_SP_SITE_ID=${siteId}`);
    console.log(`VITE_SP_DRIVE_ID=${driveId}`);
    console.log(`VITE_SP_ITEM_ID=${itemId}`);
    
  } catch (error) {
    console.error('❌ Error:', error);
  }
}

// Запуск функции
console.log('🚀 Getting SharePoint IDs...');
getSharePointIds();
