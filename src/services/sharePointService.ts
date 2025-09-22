import { Client } from "@microsoft/microsoft-graph-client";
import { AuthenticationProvider } from "@microsoft/microsoft-graph-client";

export class SharePointService {
  private graphClient: Client;
  private workbookSessionId: string | null = null;

  constructor(authProvider: AuthenticationProvider) {
    this.graphClient = Client.initWithMiddleware({ authProvider });
  }

  // Создаем (и кэшируем) сессию Excel, чтобы избежать блокировок файла
  private async getWorkbookSessionId(): Promise<string> {
    if (this.workbookSessionId) return this.workbookSessionId;
    const siteId = await this.getSiteId();
    const fileId = await this.findExcelFile();
    const session = await this.graphClient
      .api(`/sites/${siteId}/drive/items/${fileId}/workbook/createSession`)
      .post({ persistChanges: true });
    this.workbookSessionId = session?.id;
    console.log('🧩 Workbook session created:', this.workbookSessionId);
    return this.workbookSessionId!;
  }

  private resetWorkbookSession() {
    console.log('🔁 Resetting workbook session id');
    this.workbookSessionId = null;
  }

  // Public wrapper to allow UI to manually reset the Excel session
  public resetExcelSession(): void {
    this.resetWorkbookSession();
  }

  private isWorkbookLockedError(error: any): boolean {
    const code = (error?.code || error?.body?.error?.code || '').toString().toLowerCase();
    const status = error?.statusCode || error?.status;
    return code.includes('itemlocked') || code.includes('workbookbusy') || status === 423 || status === 409;
  }

  private isSessionInvalidError(error: any): boolean {
    const code = (error?.code || error?.body?.error?.code || '').toString().toLowerCase();
    const status = error?.statusCode || error?.status;
    // Common patterns when workbook-session-id is stale/invalid
    return (
      code.includes('invalidsession') ||
      code.includes('sessionnotfound') ||
      code.includes('sessionexpired') ||
      (code.includes('session') && (status === 401 || status === 404))
    );
  }

  private isTransientNetworkError(error: any): boolean {
    const status = error?.statusCode || error?.status;
    return status === 408 || status === 502 || status === 503 || status === 504;
  }

  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  // Retry helpers
  private isRateLimitError(error: any): boolean {
    const status = error?.statusCode || error?.status;
    const code = (error?.code || error?.body?.error?.code || '').toString().toLowerCase();
    return status === 429 || code.includes('ratelimit');
  }

  private isRetryableError(error: any): boolean {
    return this.isWorkbookLockedError(error) || this.isSessionInvalidError(error) || this.isTransientNetworkError(error) || this.isRateLimitError(error);
  }

  private backoffDelay(attempt: number, baseMs: number = 300, capMs: number = 5000): number {
    const jitterFactor = 0.5 + Math.random(); // 0.5..1.5
    const ms = Math.floor(baseMs * attempt * jitterFactor);
    return Math.min(ms, capMs);
  }

  // A1 helpers
  private colLettersToIndex(letters: string): number {
    let idx = 0;
    for (const ch of letters.toUpperCase()) {
      if (ch < 'A' || ch > 'Z') break;
      idx = idx * 26 + (ch.charCodeAt(0) - 64);
    }
    return idx; // 1-based
  }

  private indexToColLetters(index: number): string {
    let n = index; // 1-based
    let s = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  private parseA1Address(address: string): { sheet: string; startCol: number; startRow: number; endCol: number; endRow: number } | null {
    // Examples: "Sheet1!B2:Z200", "tubing!A1:W999", may contain quotes in sheet names
    try {
      const [sheetPart, rangePart] = address.split('!');
      if (!rangePart) return null;
      const [start, end] = rangePart.split(':');
      const m1 = start.match(/([A-Za-z]+)(\d+)/);
      const m2 = end.match(/([A-Za-z]+)(\d+)/);
      if (!m1 || !m2) return null;
      return {
        sheet: sheetPart.replace(/^'+|'+$/g, ''),
        startCol: this.colLettersToIndex(m1[1]),
        startRow: parseInt(m1[2], 10),
        endCol: this.colLettersToIndex(m2[1]),
        endRow: parseInt(m2[2], 10),
      };
    } catch {
      return null;
    }
  }

  private async getUsedRangeInfo(worksheetName: string): Promise<{ address: string; values: any[][]; meta: { startCol: number; startRow: number; endCol: number; endRow: number } } | null> {
    try {
      const siteId = await this.getSiteId();
      const fileId = await this.findExcelFile();
      let attempt = 1;
      const maxAttempts = 5;
      while (attempt <= maxAttempts) {
        try {
          const sessionId = await this.getWorkbookSessionId();
          const res = await this.graphClient
            .api(`/sites/${siteId}/drive/items/${fileId}/workbook/worksheets('${worksheetName}')/usedRange(valuesOnly=true)`) // valuesOnly avoids formatting-only cells
            .header('workbook-session-id', sessionId)
            .get();
          const address = res?.address || res?.addressLocal || '';
          const meta = this.parseA1Address(address);
          if (!meta) {
            console.warn('⚠️ Failed to parse usedRange address:', address);
            return { address, values: res?.values || [], meta: { startCol: 1, startRow: 1, endCol: (res?.values?.[0]?.length || 1), endRow: (res?.values?.length || 1) } };
          }
          return { address, values: res?.values || [], meta };
        } catch (err) {
          const status = err?.statusCode || err?.status;
          const code = (err?.code || err?.body?.error?.code || '').toString();
          const requestId = err?.response?.headers?.get?.('request-id') || err?.requestId;
          if (this.isRetryableError(err)) {
            console.warn(`⚠️ usedRange fetch failed (status ${status}, code ${code}, attempt ${attempt}/5, session reset). requestId=${requestId}`);
            this.resetWorkbookSession();
            await this.delay(this.backoffDelay(attempt));
            attempt++;
            continue;
          }
          throw err;
        }
      }
      return null;
    } catch (err) {
      console.error('❌ Error fetching usedRange info:', err);
      return null;
    }
  }

  // Проверить подключение к SharePoint
  async testConnection(): Promise<void> {
    const siteId = await this.getSiteId();
    await this.graphClient.api(`/sites/${siteId}`).get();
  }

  // Temporary test function to verify site access and drive names
  async testSiteAccess(): Promise<void> {
    try {
      console.log('=== SharePoint Access Test ===');
      
      // Test 1: Basic site access
      console.log('1. Testing site access...');
      try {
        const siteId = await this.getSiteId();
        const fileId = await this.findExcelFile();
        const site = await this.graphClient.api(`/sites/${siteId}`).get();
        console.log("✅ Site accessible:", {
          id: site.id,
          name: site.displayName,
          url: site.webUrl
        });

        // Test 2: List available drives
        console.log('2. Testing drives access...');
        try {
          const drives = await this.graphClient.api(`/sites/${siteId}/drives?$select=id,name,driveType`).get();
          console.log("✅ Drives found:", drives.value.map((d: any) => ({ 
            id: d.id, 
            name: d.name, 
            type: d.driveType 
          })));

          // Test 3: Test root folder access for each drive
          console.log('3. Testing root folder access...');
          for (const drive of drives.value) {
            try {
              const rootItems = await this.graphClient
                .api(`/sites/${siteId}/drives/${drive.id}/root/children?$select=name,folder,file&$top=10`)
                .get();
              console.log(`✅ Drive "${drive.name}" root contents (${rootItems.value.length} items):`, 
                rootItems.value.map((item: any) => ({
                  name: item.name,
                  type: item.folder ? 'folder' : 'file'
                }))
              );
            } catch (driveError) {
              console.error(`❌ Cannot access drive "${drive.name}":`, driveError);
            }
          }

          // Test 4: Try to access UPLOADS folder specifically
          console.log('4. Testing UPLOADS folder access...');
          try {
            const uploadsItems = await this.graphClient
              .api(`/sites/${siteId}/drive/items/${fileId}/children?$select=name,file`)
              .get();
            console.log("✅ UPLOADS folder contents:", uploadsItems.value.map((item: any) => item.name));
          } catch (uploadsError) {
            console.error("❌ UPLOADS folder access failed:", uploadsError);
          }

        } catch (drivesError) {
          console.error("❌ Cannot access drives:", drivesError);
        }
        
      } catch (siteError) {
        console.error("❌ Site access failed:", siteError);
        throw siteError;
      }
      
      console.log('=== Test Complete ===');
      
    } catch (error) {
      console.error('❌ SharePoint test failed:', error);
      throw error;
    }
  }

  // Получить список клиентов из SharePoint Excel файла
  async getClients(): Promise<string[]> {
    try {
      console.log('🔄 getClients() - calling getClientsFromExcel()...');
      return await this.getClientsFromExcel();
    } catch (error) {
      console.error('❌ Error in getClients():', error);
      return ['Dunga', 'KenSary', 'Tasbulat']; // Fallback data
    }
  }

  // Получить Work Orders для клиента из Excel файла
  async getWorkOrdersByClient(client: string): Promise<string[]> {
    try {
      console.log(`🔄 getWorkOrdersByClient() - calling getWorkOrdersByClient() for: ${client}`);

      // Получаем данные из листа 'wo'
      const data = await this.getExcelData('wo');
      console.log(`📊 Work orders data from Excel:`, data);

      if (!data || data.length === 0) {
        console.log('❌ No work orders data found');
        return [];
      }

      // Предполагаем, что первая строка - заголовки
      const headers = data[0];
      console.log('📋 Headers:', headers);

      // Найдем индексы нужных колонок
      const clientIndex = headers.findIndex((header: string) =>
        header && header.toLowerCase().includes('client')
      );
      const woIndex = headers.findIndex((header: string) =>
        header && (header.toLowerCase().includes('wo') || header.toLowerCase().includes('work'))
      );

      console.log(`📍 Client column index: ${clientIndex}, WO column index: ${woIndex}`);

      if (clientIndex === -1 || woIndex === -1) {
        console.log('❌ Could not find Client or WO columns');
        return [];
      }

      // Фильтруем строки по клиенту и извлекаем номера WO
      const workOrders = data.slice(1)
        .filter(row => row[clientIndex] === client)
        .map(row => row[woIndex])
        .filter(wo => wo && wo.toString().trim());

      console.log(`✅ Found ${workOrders.length} work orders for client ${client}:`, workOrders);
      return workOrders;
    } catch (error) {
      console.error('❌ Error getting work orders by client:', error);
      return [];
    }
  }

  // Создать новый Work Order в Excel файле
  async createWorkOrder(data: any): Promise<boolean> {
    try {
      console.log('🔄 Creating work order in Excel:', data);
      
      // Получить текущие данные из листа 'wo'
      const currentData = await this.getExcelData('wo');
      console.log('📊 Current work orders data:', currentData);
      
      if (!currentData || currentData.length === 0) {
        console.log('❌ No existing data found in work orders sheet');
        return false;
      }
      
      // Найти правильную позицию для вставки после последнего work order этого клиента
      const insertPosition = this.findClientInsertPosition(currentData, data.client);
      console.log(`📍 Adding work order for client ${data.client} at position: ${insertPosition}`);
      
      // Подготовить новую строку данных в том же порядке, что и заголовки
      const headers = currentData[0];
      console.log('📋 Headers:', headers);
      
      // Отладка: показать все данные формы
      console.log('🔍 DEBUG: Form data received:', data);
      console.log('🔍 DEBUG: data.diameter =', data.diameter);
      console.log('🔍 DEBUG: data.wo_date =', data.wo_date);
      
      // Создаем массив данных в порядке заголовков
      const newRowData = headers.map((header: string, index: number) => {
        const headerStr = header ? header.toString().trim() : '';
        console.log(`🔍 Column ${index}: "${headerStr}"`);
        
        // Динамическое сопоставление по названию заголовка
        const headerLower = headerStr.toLowerCase();
        
        if (headerLower.includes('wo') && !headerLower.includes('date')) {
          console.log(`   ✅ WO_No column: ${data.wo_no}`);
          return data.wo_no;
        }
        if (headerLower.includes('client')) {
          console.log(`   ✅ Client column: ${data.client}`);
          return data.client;
        }
        if (headerLower.includes('type')) {
          console.log(`   ✅ Type column: ${data.type}`);
          return data.type;
        }
        if (headerLower.includes('diameter') || headerLower.includes('диаметр')) {
          console.log(`   ✅ Diameter column: ${data.diameter}`);
          return data.diameter;
        }
        if (headerLower.includes('coupling')) {
          console.log(`   ✅ Coupling column: ${data.coupling_replace}`);
          return data.coupling_replace;
        }
        if (headerLower.includes('date')) {
          console.log(`   ✅ Date column: ${data.wo_date}`);
          return data.wo_date;
        }
        if (headerLower.includes('transport')) {
          console.log(`   ✅ Transport column: ${data.transport}`);
          return data.transport;
        }
        if (headerLower.includes('key')) {
          console.log(`   ✅ Key column: ${data.key_col}`);
          return data.key_col;
        }
        if (headerLower.includes('payer')) {
          console.log(`   ✅ Payer column: ${data.payer}`);
          return data.payer;
        }
        if (headerLower.includes('qty') || headerLower.includes('quantity')) {
          console.log(`   ✅ Quantity column: ${data.planned_qty}`);
          return data.planned_qty;
        }
        
        console.log(`   ❌ Unknown column "${headerStr}" - leaving empty`);
        return '';
      });
      
      console.log('📝 New row data:', newRowData);
      
      // Вставить новую строку в правильную позицию
      const success = await this.insertWorkOrderAtPosition(currentData, newRowData, insertPosition, data.client);
      
      if (success) {
        console.log('✅ Work order successfully added to Excel');
      } else {
        console.log('❌ Failed to add work order to Excel');
      }
      
      return success;
    } catch (error) {
      console.error('❌ Error creating work order in Excel:', error);
      return false;
    }
  }

  // Создать запись в Tubing Registry с правильным порядком вставки
  async createTubingRecord(data: any): Promise<boolean> {
    try {
      const siteId = await this.getSiteId();
      const lists = await this.graphClient
        .api(`/sites/${siteId}/lists`)
        .filter("displayName eq 'TubingRegistry'")
        .get();

      if (lists.value.length > 0) {
        const listId = lists.value[0].id;
        
        // Найти позицию для вставки (после последнего батча этого клиента)
        const insertPosition = await this.findInsertPosition(siteId, listId, data.client, data.wo_no);
        
        await this.graphClient
          .api(`/sites/${siteId}/lists/${listId}/items`)
          .post({
            fields: {
              Client: data.client,
              WO_No: data.wo_no,
              Batch: data.batch,
              Diameter: data.diameter,
              Qty: parseInt(data.qty),
              Pipe_From: parseInt(data.pipe_from),
              Pipe_To: parseInt(data.pipe_to),
              Class_1: data.class_1,
              Class_2: data.class_2,
              Class_3: data.class_3,
              Repair: data.repair,
              Scrap: data.scrap,
              Start_Date: data.start_date,
              End_Date: data.end_date,
              Rattling_Qty: parseInt(data.rattling_qty) || 0,
              External_Qty: parseInt(data.external_qty) || 0,
              Hydro_Qty: parseInt(data.hydro_qty) || 0,
              MPI_Qty: parseInt(data.mpi_qty) || 0,
              Drift_Qty: parseInt(data.drift_qty) || 0,
              EMI_Qty: parseInt(data.emi_qty) || 0,
              Marking_Qty: parseInt(data.marking_qty) || 0,
              Act_No_Oper: data.act_no_oper,
              Act_Date: data.act_date,
              SortOrder: insertPosition // Поле для правильной сортировки
            }
          });
        return true;
      }
      return false;
    } catch (error) {
      console.error('Error creating tubing record:', error);
      return false;
    }
  }

  // Найти правильную позицию для вставки Work Order после последнего WO этого клиента
  private findClientInsertPosition(currentData: any[][], client: string): number {
    if (!currentData || currentData.length <= 1) {
      return 2; // Если нет данных, вставляем во вторую строку (после заголовков)
    }

    const headers = currentData[0];
    const clientIndex = headers.findIndex((header: string) => 
      header && header.toLowerCase().includes('client')
    );

    if (clientIndex === -1) {
      console.log('❌ Client column not found, appending to end');
      return currentData.length + 1;
    }

    // Найти последнюю строку с этим клиентом
    let lastClientRow = -1;
    for (let i = currentData.length - 1; i >= 1; i--) { // Начинаем с конца, пропускаем заголовки
      if (currentData[i][clientIndex] === client) {
        lastClientRow = i;
        break;
      }
    }

    if (lastClientRow === -1) {
      // Если это первый work order для этого клиента, найти где должен быть этот клиент
      // Вставляем в конец для простоты, но можно добавить логику сортировки по алфавиту
      console.log(`📍 First work order for client ${client}, adding to end`);
      return currentData.length + 1;
    }

    // Вставляем после последнего work order этого клиента
    console.log(`📍 Found last work order for client ${client} at row ${lastClientRow + 1}, inserting after`);
    return lastClientRow + 2; // +2 потому что Excel строки начинаются с 1, а массив с 0
  }

  // Вставить Work Order в определенную позицию, сдвинув остальные строки
  private async insertWorkOrderAtPosition(currentData: any[][], newRowData: any[], insertPosition: number, client: string): Promise<boolean> {
    try {
      const headers = currentData[0];
      // Проверяем что newRowData соответствует количеству колонок
      console.log(`🔍 Headers length: ${headers.length}, newRowData length: ${newRowData.length}`);
      console.log(`🔍 NewRowData (raw):`, newRowData);

      // Убедимся что newRowData имеет правильную длину
      while (newRowData.length < headers.length) {
        newRowData.push('');
      }
      if (newRowData.length > headers.length) {
        newRowData = newRowData.slice(0, headers.length);
      }

      // Получаем usedRange, чтобы учитывать смещение по колонкам/строкам (как в tubing)
      const usedInfo = await this.getUsedRangeInfo('wo');
      const startColIdx = usedInfo?.meta?.startCol ?? 1; // 1-based
      const startRow = usedInfo?.meta?.startRow ?? 1; // 1-based
      // Используем полную ширину usedRange, т.к. она может быть шире, чем длина headers
      const usedWidth = usedInfo?.meta ? (usedInfo.meta.endCol - usedInfo.meta.startCol + 1) : headers.length;
      // Абсолютный номер строки в Excel для позиции вставки (включая смещение usedRange)
      const absoluteInsertRow = startRow + (insertPosition - 1);

      // Если вставляем в конец, просто добавляем в следующую пустую строку
      if (insertPosition > currentData.length) {
        const appendRow = startRow + currentData.length; // следующая пустая строка после usedRange
        const range = `${startColLetters}${appendRow}:${endColLetters}${appendRow}`;
        console.log(`📍 Appending work order to end at range: ${range}`);
        // Нормализуем ширину строки под usedRange
        const normalizeRow = (row: any[]) => {
          const r = Array.isArray(row) ? [...row] : [];
          while (r.length < usedWidth) r.push('');
          if (r.length > usedWidth) r.length = usedWidth;
          return r.map(cell => (cell === null || cell === undefined || cell === '') ? '' : String(cell).trim());
        };
        const cleanedData = [normalizeRow(newRowData)];
        const ok = await this.writeExcelData('wo', range, cleanedData);
        if (ok) {
          console.log(`✅ Work order appended successfully!`);
          localStorage.removeItem('sharepoint_cached_wo');
          localStorage.removeItem('sharepoint_cache_timestamp_wo');
        } else {
          console.log(`❌ Failed to append work order`);
        }
        return ok;
      }

      // Иначе нужно сдвинуть существующие строки вниз на одну позицию
      console.log(`📍 Inserting at logical position ${insertPosition} (absolute row ${absoluteInsertRow}). Will shift rows down.`);
      const rowsToShift = currentData.slice(insertPosition - 1); // данные начиная с строки вставки

      // 1) Вставляем физическую пустую строку на нужном месте (Excel сам сдвигает вниз)
      const newRowRange = `${startColLetters}${absoluteInsertRow}:${endColLetters}${absoluteInsertRow}`;
      const rowAddress = `${absoluteInsertRow}:${absoluteInsertRow}`; // вставка целой строки
      console.log(`➕ Inserting full row at: ${rowAddress}`);
      const inserted = await this.insertRowAt('wo', rowAddress);
      if (!inserted) {
        console.warn('⚠️ Row insert failed, fallback to writing directly (may risk overlap)');
      }

      // 2) Записываем новую строку в освободившийся диапазон
      console.log(`📝 Writing new work order row to range: ${newRowRange}`);
      const normalizeRow = (row: any[]) => {
        const r = Array.isArray(row) ? [...row] : [];
        while (r.length < usedWidth) r.push('');
        if (r.length > usedWidth) r.length = usedWidth;
        return r.map(cell => (cell === null || cell === undefined || cell === '') ? '' : String(cell).trim());
      };
      const cleanedNewRow = [normalizeRow(newRowData)];
      const writeNewRowSuccess = await this.writeExcelData('wo', newRowRange, cleanedNewRow);
      if (!writeNewRowSuccess) {
        console.log('❌ Failed to write new work order row');
        return false;
      }

      console.log(`✅ Successfully inserted work order for client ${client} at absolute row ${absoluteInsertRow}`);
      // Очистим кэш, чтобы форсировать обновление данных
      localStorage.removeItem('sharepoint_cached_wo');
      localStorage.removeItem('sharepoint_cache_timestamp_wo');
      return true;

    } catch (error) {
      console.error('❌ Error inserting work order at position:', error);
      return false;
    }
  }

  // Найти правильную позицию для вставки Tubing записи после последнего батча этого клиента/WO
  private findTubingInsertPosition(currentData: any[][], client: string, woNo: string): number {
    if (!currentData || currentData.length <= 1) {
      return 2; // Если нет данных, вставляем во вторую строку (после заголовков)
    }

    const headers = currentData[0];
    const normalize = (value: any) =>
      value === null || value === undefined
        ? ''
        : String(value).trim().toLowerCase();
    const targetClient = normalize(client);
    const targetWo = normalize(woNo);

    const clientIndex = headers.findIndex((header: string) =>
      header && String(header).toLowerCase().includes('client')
    );
    const woIndex = headers.findIndex((header: string) =>
      header && String(header).toLowerCase().includes('wo')
    );

    if (clientIndex === -1 || woIndex === -1) {
      console.log('❌ Client or WO column not found, appending to end');
      return currentData.length + 1;
    }

    // Найти последнюю строку с этим клиентом и WO
    let lastClientWoRow = -1;
    for (let i = currentData.length - 1; i >= 1; i--) { // Начинаем с конца, пропускаем заголовки
      const rowClient = normalize(currentData[i][clientIndex]);
      const rowWo = normalize(currentData[i][woIndex]);
      if (!rowClient && !rowWo) continue; // пропускаем полностью пустые строки
      if (rowClient === targetClient && rowWo === targetWo) {
        lastClientWoRow = i;
        break;
      }
    }

    if (lastClientWoRow === -1) {
      // Если это первая запись для этого клиента/WO, найти где должен быть этот клиент
      // Найти последнюю запись этого клиента (любого WO)
      let lastClientRow = -1;
      for (let i = currentData.length - 1; i >= 1; i--) {
        const rowClient = normalize(currentData[i][clientIndex]);
        if (!rowClient) continue;
        if (rowClient === targetClient) {
          lastClientRow = i;
          break;
        }
      }
      
      if (lastClientRow === -1) {
        // Если это первая запись клиента вообще, добавляем в конец
        console.log(`📍 First tubing record for client ${client}, adding to end`);
        return currentData.length + 1;
      } else {
        // Вставляем после последней записи этого клиента
        console.log(`📍 First WO ${woNo} for client ${client}, inserting after last client record at row ${lastClientRow + 1}`);
        return lastClientRow + 2;
      }
    }

    // Вставляем после последней записи этого клиента/WO
    console.log(`📍 Found last tubing record for client ${client}, WO ${woNo} at row ${lastClientWoRow + 1}, inserting after`);
    return lastClientWoRow + 2; // +2 потому что Excel строки начинаются с 1, а массив с 0
  }

  // Вставить Tubing запись в определенную позицию, сдвинув остальные строки
  private async insertTubingAtPosition(currentData: any[][], newRowData: any[], insertPosition: number, client: string, woNo: string): Promise<boolean> {
    try {
      const headers = currentData[0];
      // Проверяем что newRowData соответствует количеству колонок
      console.log(`🔍 Headers length: ${headers.length}, newRowData length: ${newRowData.length}`);
      console.log(`🔍 NewRowData (raw):`, newRowData);

      // Убедимся что newRowData имеет правильную длину
      while (newRowData.length < headers.length) {
        newRowData.push('');
      }
      if (newRowData.length > headers.length) {
        newRowData = newRowData.slice(0, headers.length);
      }

      // Получаем usedRange, чтобы учитывать смещение по колонкам/строкам
      const usedInfo = await this.getUsedRangeInfo('tubing');
      const startColIdx = usedInfo?.meta?.startCol ?? 1; // 1-based
      const startRow = usedInfo?.meta?.startRow ?? 1; // 1-based
      // Используем полную ширину usedRange (она может быть шире headers)
      // Абсолютный номер строки в Excel для позиции вставки (включая смещение usedRange)
      const absoluteInsertRow = startRow + (insertPosition - 1);

      let targetRowNumber = absoluteInsertRow;
      // Если вставляем в конец, используем следующую пустую строку
      if (insertPosition > currentData.length) {
        targetRowNumber = startRow + currentData.length;
        console.log(`📍 Appending tubing record to end at row ${targetRowNumber}`);
      } else {
        console.log(`📍 Inserting at logical position ${insertPosition} (absolute row ${absoluteInsertRow}). Will shift rows down.`);
        const rowAddress = `${absoluteInsertRow}:${absoluteInsertRow}`;
        console.log(`➕ Inserting full row at: ${rowAddress}`);
        const inserted = await this.insertRowAt('tubing', rowAddress);
        if (!inserted) {
          console.warn('⚠️ Row insert failed, fallback to writing directly (may risk overlap)');
        }
      }

      const shouldWriteCell = (value: any) => {
        if (value === null || value === undefined) return false;
        if (typeof value === 'string') {
          return value.trim() !== '';
        }
        return true;
      };

      const updates = newRowData
        .map((value, idx) => ({ value, idx }))
        .filter(update => shouldWriteCell(update.value));

      if (updates.length === 0) {
        console.log('ℹ️ No explicit values to write for new tubing row, preserving formulas.');
      }

      const writeGroups: { startIdx: number; endIdx: number; values: any[] }[] = [];
      updates.forEach(update => {
        const lastGroup = writeGroups[writeGroups.length - 1];
        if (lastGroup && update.idx === lastGroup.endIdx + 1) {
          lastGroup.endIdx = update.idx;
          lastGroup.values.push(update.value);
        } else {
          writeGroups.push({ startIdx: update.idx, endIdx: update.idx, values: [update.value] });
        }
      });

      for (const group of writeGroups) {
        const startCol = startColIdx + group.startIdx;
        const endCol = startColIdx + group.endIdx;
        const range = `${this.indexToColLetters(startCol)}${targetRowNumber}:${this.indexToColLetters(endCol)}${targetRowNumber}`;
        console.log(`📝 Writing tubing data to range ${range}`);
        const ok = await this.writeExcelData('tubing', range, [group.values]);
        if (!ok) {
          console.log(`❌ Failed to write tubing data to range ${range}`);
          return false;
        }
      }

      console.log(`✅ Successfully inserted tubing for client ${client}, WO ${woNo} at absolute row ${targetRowNumber}`);
      localStorage.removeItem('sharepoint_cached_tubing');
      localStorage.removeItem('sharepoint_cache_timestamp_tubing');
      return true;
      
    } catch (error) {
      console.error('❌ Error inserting tubing record at position:', error);
      return false;
    }
  }

  // Найти правильную позицию для вставки записи
  private async findInsertPosition(siteId: string, listId: string, client: string, woNo: string): Promise<number> {
    try {
      // Получить все записи этого клиента и WO
      const clientRecords = await this.graphClient
        .api(`/sites/${siteId}/lists/${listId}/items`)
        .expand('fields')
        .filter(`fields/Client eq '${client}' and fields/WO_No eq '${woNo}'`)
        .orderby('fields/SortOrder desc')
        .top(1)
        .get();

      if (clientRecords.value.length > 0) {
        return clientRecords.value[0].fields.SortOrder + 1;
      }

      // Если нет записей для этого клиента, найти последнюю запись клиента вообще
      const lastClientRecord = await this.graphClient
        .api(`/sites/${siteId}/lists/${listId}/items`)
        .expand('fields')
        .filter(`fields/Client eq '${client}'`)
        .orderby('fields/SortOrder desc')
        .top(1)
        .get();

      if (lastClientRecord.value.length > 0) {
        return lastClientRecord.value[0].fields.SortOrder + 1;
      }

      // Если это первая запись клиента, найти общую последнюю позицию
      const allRecords = await this.graphClient
        .api(`/sites/${siteId}/lists/${listId}/items`)
        .expand('fields')
        .orderby('fields/SortOrder desc')
        .top(1)
        .get();

      return allRecords.value.length > 0 ? allRecords.value[0].fields.SortOrder + 100 : 1000;
    } catch (error) {
      console.error('Error finding insert position:', error);
      return Date.now(); // Fallback to timestamp
    }
  }

  // Получить список листов в Excel файле
  async getWorksheetNames(): Promise<string[]> {
    try {
      const siteId = await this.getSiteId();
      const fileId = await this.findExcelFile();
      console.log('Getting worksheet names for file ID:', fileId);

      let attempt = 1;
      const maxAttempts = 5;
      while (attempt <= maxAttempts) {
        try {
          const sessionId = await this.getWorkbookSessionId();
          const worksheets = await this.graphClient
            .api(`/sites/${siteId}/drive/items/${fileId}/workbook/worksheets`)
            .header('workbook-session-id', sessionId)
            .get();

          const sheetNames = worksheets.value.map((ws: any) => ws.name);
          console.log('Available worksheets:', sheetNames);
          console.log('🔍 Looking for client sheet in:', sheetNames);
          return worksheets.value.map((ws: any) => ws.name);
        } catch (err) {
          const status = err?.statusCode || err?.status;
          const code = (err?.code || err?.body?.error?.code || '').toString();
          const requestId = err?.response?.headers?.get?.('request-id') || err?.requestId;
          if (this.isRetryableError(err)) {
            console.warn(`⚠️ Worksheets fetch failed (status ${status}, code ${code}, attempt ${attempt}/${maxAttempts}). Resetting session and retrying... requestId=${requestId}`);
            this.resetWorkbookSession();
            await this.delay(this.backoffDelay(attempt));
            attempt++;
            continue;
          }
          throw err;
        }
      }
      return [];
    } catch (error) {
      console.error('Error getting worksheet names:', error);
      return [];
    }
  }

  // Получить данные из Excel файла pipe_inspection.xlsm
  async getExcelData(worksheetName: string, range?: string): Promise<any[]> {
    try {
      console.log(`Getting Excel data for worksheet: ${worksheetName}`);
      
      // Сначала получим список доступных листов
      const availableSheets = await this.getWorksheetNames();
      console.log('Available sheets:', availableSheets);
      
      // Если запрашиваемый лист не существует, попробуем найти похожий
      let actualSheetName = worksheetName;
      if (!availableSheets.includes(worksheetName)) {
        console.log(`Sheet '${worksheetName}' not found, looking for alternatives...`);
        
        // Попробуем найти лист по частичному совпадению
        const foundSheet = availableSheets.find(sheet => 
          sheet.toLowerCase().includes(worksheetName.toLowerCase()) ||
          worksheetName.toLowerCase().includes(sheet.toLowerCase())
        );
        
        if (foundSheet) {
          actualSheetName = foundSheet;
          console.log(`Using alternative sheet: ${actualSheetName}`);
        } else {
          console.log(`❌ No matching sheet found for '${worksheetName}'. Available sheets:`, availableSheets);
        console.log('🔍 Exact sheet names:', availableSheets.map(s => `'${s}'`));
          return [];
        }
      }

      const siteId = await this.getSiteId();
      
      const fileId = await this.findExcelFile();
      
      // Получить данные из указанного листа
      let worksheetApiPath;
      if (range) {
        // Для конкретного диапазона используем правильный синтаксис
        worksheetApiPath = `/sites/${siteId}/drive/items/${fileId}/workbook/worksheets('${actualSheetName}')/range(address='${range}')`;
      } else {
        // Для всех данных используем usedRange
        worksheetApiPath = `/sites/${siteId}/drive/items/${fileId}/workbook/worksheets('${actualSheetName}')/usedRange(valuesOnly=true)`;
      }
      
      console.log(`Getting worksheet data: ${worksheetApiPath}`);
      
      let attempt = 1;
      const maxAttempts = 5;
      while (attempt <= maxAttempts) {
        try {
          const sessionId = await this.getWorkbookSessionId();
          const worksheetData = await this.graphClient
            .api(worksheetApiPath)
            .header('workbook-session-id', sessionId)
            .get();
          console.log('Worksheet data:', worksheetData);
          return worksheetData.values || [];
        } catch (err) {
          const status = err?.statusCode || err?.status;
          const code = (err?.code || err?.body?.error?.code || '').toString();
          const requestId = err?.response?.headers?.get?.('request-id') || err?.requestId;
          if (this.isRetryableError(err)) {
            console.warn(`⚠️ Worksheet read failed (status ${status}, code ${code}, attempt ${attempt}/${maxAttempts}). Resetting session and retrying... requestId=${requestId}`);
            this.resetWorkbookSession();
            await this.delay(this.backoffDelay(attempt));
            attempt++;
            continue;
          }
          throw err;
        }
      }
      return [];
    } catch (error) {
      console.error('Error reading Excel data:', error);
      return [];
    }
  }

  // Записать данные в Excel файл
  async writeExcelData(worksheetName: string, range: string, values: any[][]): Promise<boolean> {
    try {
      const siteId = await this.getSiteId();
      
      const fileId = await this.findExcelFile();

      // Resolve actual worksheet name (case-insensitive / partial match), similar to getExcelData
      let actualSheetName = worksheetName;
      try {
        const availableSheets = await this.getWorksheetNames();
        if (!availableSheets.includes(worksheetName)) {
          const foundSheet =
            availableSheets.find(s => s.toLowerCase() === worksheetName.toLowerCase()) ||
            availableSheets.find(s => s.toLowerCase().includes(worksheetName.toLowerCase()) || worksheetName.toLowerCase().includes(s.toLowerCase()));
          if (foundSheet) {
            console.log(`✳️ Resolved worksheet '${worksheetName}' to '${foundSheet}' for write`);
            actualSheetName = foundSheet;
          } else {
            console.warn(`⚠️ Worksheet '${worksheetName}' not found among:`, availableSheets);
          }
        }
      } catch (resolveErr) {
        console.warn('⚠️ Could not resolve worksheet name, proceeding with provided name:', worksheetName, resolveErr);
      }

      const doPatch = async () => {
        const sessionId = await this.getWorkbookSessionId();
        console.log('📝 Writing to Excel', { sheet: actualSheetName, range, rows: values?.length, cols: values?.[0]?.length, sessionId });
        await this.graphClient
          .api(`/sites/${siteId}/drive/items/${fileId}/workbook/worksheets('${actualSheetName}')/range(address='${range}')`)
          .header('workbook-session-id', sessionId)
          .patch({ values });
      };

      let attempt = 1;
      const maxAttempts = 5;
      while (attempt <= maxAttempts) {
        try {
          await doPatch();
          return true;
        } catch (err) {
          const status = err?.statusCode || err?.status;
          const code = (err?.code || err?.body?.error?.code || '').toString();
          const requestId = err?.response?.headers?.get?.('request-id') || err?.requestId;
          if (this.isRetryableError(err)) {
            console.warn(`⚠️ Write failed (status ${status}, code ${code}) on attempt ${attempt}/${maxAttempts}. Resetting session and retrying... requestId=${requestId}`);
            this.resetWorkbookSession();
            await this.delay(this.backoffDelay(attempt, 600, 8000));
            attempt++;
            continue;
          }
          throw err;
        }
      }
      return false;
    } catch (error) {
      console.error('Error writing Excel data:', error);
      return false;
    }
  }

  // Вставить пустую строку в указанном диапазоне (сдвиг вниз всей области)
  private async insertRowAt(worksheetName: string, range: string): Promise<boolean> {
    try {
      const siteId = await this.getSiteId();
      const fileId = await this.findExcelFile();

      // Разрешаем фактическое имя листа, как и при записи
      let actualSheetName = worksheetName;
      try {
        const availableSheets = await this.getWorksheetNames();
        if (!availableSheets.includes(worksheetName)) {
          const foundSheet =
            availableSheets.find(s => s.toLowerCase() === worksheetName.toLowerCase()) ||
            availableSheets.find(s => s.toLowerCase().includes(worksheetName.toLowerCase()) || worksheetName.toLowerCase().includes(s.toLowerCase()));
          if (foundSheet) {
            actualSheetName = foundSheet;
            console.log(`✳️ Resolved worksheet '${worksheetName}' to '${foundSheet}' for insert`);
          }
        }
      } catch (resolveErr) {
        console.warn('⚠️ Could not resolve worksheet name for insert, proceeding with provided name:', worksheetName, resolveErr);
      }

      const doInsert = async () => {
        const sessionId = await this.getWorkbookSessionId();
        console.log('➕ Inserting row', { sheet: actualSheetName, range });
        await this.graphClient
          .api(`/sites/${siteId}/drive/items/${fileId}/workbook/worksheets('${actualSheetName}')/range(address='${range}')/insert`)
          .header('workbook-session-id', sessionId)
          .post({ shift: 'Down' });
      };

      let attempt = 1;
      const maxAttempts = 5;
      while (attempt <= maxAttempts) {
        try {
          await doInsert();
          return true;
        } catch (err) {
          const status = err?.statusCode || err?.status;
          const code = (err?.code || err?.body?.error?.code || '').toString();
          const requestId = err?.response?.headers?.get?.('request-id') || err?.requestId;
          if (this.isRetryableError(err)) {
            console.warn(`⚠️ Insert row failed (status ${status}, code ${code}) on attempt ${attempt}/${maxAttempts}. Resetting session and retrying... requestId=${requestId}`);
            this.resetWorkbookSession();
            await this.delay(this.backoffDelay(attempt, 600, 8000));
            attempt++;
            continue;
          }
          throw err;
        }
      }
      return false;
    } catch (error) {
      console.error('Error inserting row:', error);
      return false;
    }
  }

  // Получить клиентов из Excel файла
  async getClientsFromExcel(): Promise<string[]> {
    try {
      console.log('🔍 Starting to load clients from Excel...');
      
      // Сначала получим все доступные листы
      const worksheets = await this.getWorksheetNames();
      console.log('📋 All available worksheets:', worksheets);
      
      // Попробуем разные варианты названий листа с клиентами
      const possibleClientSheets = ['Client', 'Clients', 'client', 'clients', 'wo'];
      let clientSheet = null;
      
      for (const sheetName of possibleClientSheets) {
        if (worksheets.includes(sheetName)) {
          clientSheet = sheetName;
          console.log(`✅ Found client sheet: '${clientSheet}'`);
          break;
        }
      }
      
      if (!clientSheet) {
        console.log('❌ No client sheet found. Trying first sheet with data...');
        clientSheet = worksheets[0]; // Используем первый лист
      }
      
      // Получаем все данные из листа 'wo' и извлекаем колонку B
      const data = await this.getExcelData(clientSheet); // Получаем все данные
      console.log(`📊 Full data from sheet '${clientSheet}':`, data);
      
      // Извлекаем только колонку B (индекс 1)
      const columnBData = data.map(row => row[1]).filter(cell => cell && cell.trim());
      console.log(`📊 Column B data from sheet '${clientSheet}':`, columnBData);
      
      // Пропускаем заголовок (первую строку), фильтруем пустые и убираем дубликаты
      const clients = [...new Set(columnBData.slice(1).filter(client => client && client.trim()))];
      console.log('🔄 Unique clients after removing duplicates:', clients);
      console.log('✅ Filtered clients:', clients);
      
      return clients;
    } catch (error) {
      console.error('❌ Error getting clients from Excel:', error);
      return ['Dunga', 'KenSary', 'Tasbulat']; // Fallback
    }
  }

  // Получить Work Orders из Excel файла
  async getWorkOrdersFromExcel(client: string): Promise<string[]> {
    try {
      const data = await this.getExcelData('WorkOrders');
      const headers = data[0];
      const clientIndex = headers.indexOf('Client');
      const woIndex = headers.indexOf('WO_No');

      return data.slice(1)
        .filter(row => row[clientIndex] === client)
        .map(row => row[woIndex])
        .filter(wo => wo);
    } catch (error) {
      console.error('Error getting work orders from Excel:', error);
      return [];
    }
  }
  // Добавить новую запись в Excel файл - ТОЧНО КАК WO FORM
  async addTubingRecordToExcel(data: any): Promise<boolean> {
    try {
      console.log('🔥 TUBING: Starting addTubingRecordToExcel with data:', data);

      // Получить текущие данные из листа 'tubing'
      console.log('📊 TUBING: Getting current tubing data...');
      let currentData = await this.getExcelData('tubing');
      console.log('📊 TUBING: Current data length:', currentData?.length);

      if (!currentData || currentData.length === 0) {
        console.warn('⚠️ TUBING: Empty data returned. Resetting session and retrying read...');
        this.resetWorkbookSession();
        await this.delay(300);
        currentData = await this.getExcelData('tubing');
        if (!currentData || currentData.length === 0) {
          console.log('❌ TUBING: No existing data found in tubing sheet after retry');
          return false;
        }
      }

      // Подготовить новую строку данных в том же порядке, что и заголовки
      const headers = currentData[0];
      console.log('📋 TUBING: Headers found:', headers);

      // Создаем массив данных в порядке заголовков (безопасно обрабатываем нестроковые заголовки)
      const newRowData = headers.map((header: string) => {
        const headerStr = header !== undefined && header !== null ? String(header) : '';
        const headerLower = headerStr.toLowerCase();
        if (headerLower.includes('client')) return data.client;
        if (headerLower.includes('wo')) return data.wo_no;
        if (headerLower.includes('batch')) return data.batch;
        if (headerLower.includes('diameter')) return data.diameter;
        if (headerLower.includes('qty') && !headerLower.includes('_')) return data.qty;
        if (headerLower.includes('pipe_from') || headerLower.includes('from')) return data.pipe_from;
        if (headerLower.includes('pipe_to') || headerLower.includes('to')) return data.pipe_to;
        if (headerLower.includes('rack')) return data.rack || '';
        if (headerLower.includes('status')) return data.status || 'Arrived';
        if (headerLower.includes('arrival')) return data.arrival_date;
        return ''; // Пустое значение для неизвестных колонок
      });

      console.log('📝 TUBING: New row data prepared:', newRowData);

      // Найдём правильную позицию вставки для сохранения группировки Клиент → WO
      const insertPosition = this.findTubingInsertPosition(currentData, data.client, data.wo_no);
      console.log(`📍 TUBING: Computed insert position (relative to usedRange): ${insertPosition}`);

      const success = await this.insertTubingAtPosition(currentData, newRowData, insertPosition, data.client, data.wo_no);

      if (success) {
        console.log('✅ TUBING: Record inserted successfully at grouped position');
      } else {
        console.log('❌ TUBING: Failed to insert tubing record');
      }

      return success;
    } catch (error) {
      console.error('❌ Error adding tubing record to Excel:', error);
      return false;
    }
  }

  async updateTubingInspectionData(data: {
    client: string;
    wo_no: string;
    batch: string;
    class_1?: string;
    class_2?: string;
    class_3?: string;
    repair?: string;
    scrap?: string | number;
    rattling_qty?: number;
    external_qty?: number;
    hydro_qty?: number;
    mpi_qty?: number;
    drift_qty?: number;
    emi_qty?: number;
    marking_qty?: number;
    rattling_scrap_qty?: number;
    external_scrap_qty?: number;
    jetting_scrap_qty?: number;
    mpi_scrap_qty?: number;
    drift_scrap_qty?: number;
    emi_scrap_qty?: number;
    start_date?: string;
    end_date?: string;
    status?: string;
  }): Promise<boolean> {
    try {
      const usedInfo = await this.getUsedRangeInfo('tubing');
      if (!usedInfo?.values?.length) {
        console.warn('⚠️ No tubing data available to update');
        return false;
      }

      const { values, meta } = usedInfo;
      const headersRow = Array.isArray(values[0]) ? (values[0] as unknown[]) : [];

      const normalize = (value: unknown) =>
        value === null || value === undefined
          ? ''
          : String(value).trim().toLowerCase();
      const simplify = (value: unknown) =>
        normalize(value).replace(/[\s_-]+/g, '');

      const findColumn = (matcher: (normalized: string, simplified: string) => boolean) =>
        headersRow.findIndex(header => matcher(normalize(header), simplify(header)));

      const clientIndex = findColumn((normalized, simplified) => normalized.includes('client') || simplified.includes('client'));
      const woIndex = findColumn((normalized, simplified) => normalized.includes('wo') || simplified.includes('workorder'));
      const batchIndex = findColumn((normalized, simplified) => normalized.includes('batch') || simplified.includes('batch'));

      if (clientIndex === -1 || woIndex === -1 || batchIndex === -1) {
        console.error('❌ Required columns (client/wo/batch) not found in tubing sheet');
        return false;
      }

      const targetClient = normalize(data.client);
      const targetWo = normalize(data.wo_no);
      const targetBatch = normalize(data.batch);

      const rowIndex = values.findIndex((row, idx) => {
        if (idx === 0) return false;
        return (
          normalize(row[clientIndex]) === targetClient &&
          normalize(row[woIndex]) === targetWo &&
          normalize(row[batchIndex]) === targetBatch
        );
      });

      if (rowIndex === -1) {
        console.warn('⚠️ Target tubing record not found for inspection update', data);
        return false;
      }

      const rowValues = Array.isArray(values[rowIndex]) ? (values[rowIndex] as unknown[]) : [];
      const targetRow = [...rowValues];
      const usedWidth = meta.endCol - meta.startCol + 1;
      while (targetRow.length < usedWidth) targetRow.push('');
      if (targetRow.length > usedWidth) targetRow.length = usedWidth;

      const applyValue = (
        predicate: (normalized: string, simplified: string) => boolean,
        value: unknown
      ) => {
        const columnIndex = findColumn(predicate);
        if (columnIndex !== -1) {
          targetRow[columnIndex] = value ?? '';
        }
      };

      applyValue(
        (normalized, simplified) =>
          normalized.includes('class 1') || normalized.includes('class_1') || simplified.includes('class1'),
        data.class_1
      );
      applyValue(
        (normalized, simplified) =>
          normalized.includes('class 2') || normalized.includes('class_2') || simplified.includes('class2'),
        data.class_2
      );
      applyValue(
        (normalized, simplified) =>
          normalized.includes('class 3') || normalized.includes('class_3') || simplified.includes('class3'),
        data.class_3
      );
      applyValue((normalized, simplified) => normalized.includes('repair') || simplified.includes('repair'), data.repair);
      applyValue((normalized, simplified) => normalized.includes('status') || simplified.includes('status'), data.status ?? 'Inspection Done');
      applyValue(
        (normalized, simplified) =>
          (normalized.includes('scrap') || simplified.includes('scrap')) && !simplified.includes('scrapqty'),
        data.scrap ?? ''
      );

      applyValue((_, simplified) => simplified.includes('rattlingqty') && !simplified.includes('scrap'), data.rattling_qty ?? '');
      applyValue((_, simplified) => simplified.includes('externalqty'), data.external_qty ?? '');
      applyValue((_, simplified) => simplified.includes('hydroqty') || simplified.includes('jettingqty'), data.hydro_qty ?? '');
      applyValue((_, simplified) => simplified.includes('mpiqty'), data.mpi_qty ?? '');
      applyValue((_, simplified) => simplified.includes('driftqty'), data.drift_qty ?? '');
      applyValue((_, simplified) => simplified.includes('emiqty'), data.emi_qty ?? '');
      applyValue((_, simplified) => simplified.includes('markingqty'), data.marking_qty ?? '');

      applyValue((_, simplified) => simplified.includes('rattlingscrapqty'), data.rattling_scrap_qty ?? '');
      applyValue((_, simplified) => simplified.includes('externalscrapqty'), data.external_scrap_qty ?? '');
      applyValue((_, simplified) => simplified.includes('jettingscrapqty'), data.jetting_scrap_qty ?? '');
      applyValue((_, simplified) => simplified.includes('mpiscrapqty'), data.mpi_scrap_qty ?? '');
      applyValue((_, simplified) => simplified.includes('driftscrapqty'), data.drift_scrap_qty ?? '');
      applyValue((_, simplified) => simplified.includes('emiscrapqty'), data.emi_scrap_qty ?? '');
      applyValue((_, simplified) => simplified.includes('startdate'), data.start_date ?? '');
      applyValue((_, simplified) => simplified.includes('enddate'), data.end_date ?? '');

      const startColLetters = this.indexToColLetters(meta.startCol);
      const endColLetters = this.indexToColLetters(meta.startCol + usedWidth - 1);
      const rowNumber = meta.startRow + rowIndex;
      const range = `${startColLetters}${rowNumber}:${endColLetters}${rowNumber}`;

      const writeSuccess = await this.writeExcelData('tubing', range, [targetRow]);
      if (writeSuccess) {
        localStorage.removeItem('sharepoint_cached_tubing');
        localStorage.removeItem('sharepoint_cache_timestamp_tubing');
      }

      return writeSuccess;
    } catch (error) {
      console.error('❌ Error updating tubing inspection data:', error);
      return false;
    }
  }

  // Получить ID сайта SharePoint из env переменных
  private async getSiteId(): Promise<string> {
    const SITE_ID = import.meta.env.VITE_SP_SITE_ID as string;
    
    if (SITE_ID) {
      console.log('Using site ID from env:', SITE_ID);
      return SITE_ID;
    }
    
    // Fallback - получить site ID через API
    try {
      console.log('Getting SharePoint site ID from API...');
      const site = await this.graphClient
        .api('/sites/kzprimeestate.sharepoint.com:/sites/pipe-inspection?$select=id')
        .get();
      console.log('SharePoint site found:', site);
      return site.id;
    } catch (error) {
      console.error('Access denied to SharePoint site. Check user membership and Graph delegated permissions.', error);
      throw new Error('Access denied to SharePoint site. Check user membership and Graph delegated permissions.');
    }
  }

  // Найти Excel файл в SharePoint используя env переменные
  private async findExcelFile(): Promise<string> {
    try {
      const siteId = await this.getSiteId();
      const filePath = import.meta.env.VITE_SP_FILE_PATH as string;
      
      console.log('Searching for Excel file with site ID:', siteId);
      console.log('Using file path from env:', filePath);

      if (filePath) {
        // Используем прямой путь к файлу из env
        try {
          const item = await this.graphClient
            .api(`/sites/${siteId}/drive/root:${filePath}`)
            .get();
          console.log('Found Excel file by direct path:', item);
          return item.id;
        } catch (directPathError) {
          console.warn('Primary path lookup failed:', directPathError);
          // Fallback: if path likely missing default library prefix, try "/Shared Documents" prefix
          try {
            if (!filePath.startsWith('/Shared Documents') && filePath.startsWith('/')) {
              const altPath = `/Shared Documents${filePath}`;
              console.log(`Trying fallback path: ${altPath}`);
              const itemAlt = await this.graphClient
                .api(`/sites/${siteId}/drive/root:${altPath}`)
                .get();
              console.log('Found Excel file by fallback path:', itemAlt);
              return itemAlt.id;
            }
          } catch (altErr) {
            console.error('Fallback path lookup failed:', altErr);
          }
          // If both attempts failed, propagate a clear error
          throw new Error('Access denied or file not found. Verify VITE_SP_FILE_PATH and user permissions (Edit) on the site.');
        }
      }

      throw new Error('File path not configured in environment variables');
    } catch (error) {
      console.error('Error finding Excel file:', error);
      throw error;
    }
  }
}
