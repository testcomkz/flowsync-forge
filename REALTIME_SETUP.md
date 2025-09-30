# 🚀 Real-Time Setup Guide

## 1. Azure AD - Создать Client Secret

1. Открой [Azure Portal](https://portal.azure.com)
2. Azure Active Directory → App registrations
3. Найди твой app: `7c4aabf5-ed57-42e8-a828-506b2f6378da`
4. **Certificates & secrets** → New client secret
5. Description: `Vercel Serverless`
6. Expires: `24 months`
7. **Copy value** (показывается только один раз!)

## 2. Azure AD - API Permissions

1. В том же app → **API permissions**
2. Add permission → Microsoft Graph → **Application permissions**
3. Добавь:
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`
4. **Grant admin consent** (нужны права админа)

## 3. Supabase - Запустить миграцию

```bash
# В Supabase Dashboard → SQL Editor
# Скопируй содержимое файла supabase/migrations/001_realtime_tables.sql
# Запусти SQL
```

Или через CLI:
```bash
npx supabase db push
```

## 4. Supabase - Включить Realtime

1. Supabase Dashboard → Database → Replication
2. Найди таблицу `sharepoint_updates`
3. Включи **Realtime**

## 5. Local Test - ngrok

```bash
# Запусти ngrok
ngrok http 3000

# Скопируй HTTPS URL (например: https://abc123.ngrok.io)
```

## 6. Vercel - Environment Variables

```bash
AZURE_CLIENT_ID=7c4aabf5-ed57-42e8-a828-506b2f6378da
AZURE_TENANT_ID=c2b50854-0dc6-45bb-b58e-609a4ced8b6e
AZURE_CLIENT_SECRET=<твой-client-secret>
WEBHOOK_SECRET=<любая-случайная-строка>
WEBHOOK_URL=https://your-app.vercel.app
SUPABASE_URL=https://hkndhnctlctvybfftyqu.supabase.co
SUPABASE_SERVICE_KEY=<service-role-key>
SHAREPOINT_SITE_ID=kzprimeestate.sharepoint.com,9f482633-8093-471a-a0e4-c6353a265373,df166f91-e314-4b9b-87a7-cbd2e26f2c1c
SHAREPOINT_FILE_ID=<получить-через-api>
```

## 7. Получить File ID

```bash
# Запусти локально
npm run dev

# В браузере консоль:
# Подключись к SharePoint
# В консоли найди: "File ID: ..."
```

Или через API:
```bash
curl -X GET "https://graph.microsoft.com/v1.0/sites/kzprimeestate.sharepoint.com,9f482633-8093-471a-a0e4-c6353a265373,df166f91-e314-4b9b-87a7-cbd2e26f2c1c/drive/root:/UPLOADS/pipe_inspection.xlsm" \
  -H "Authorization: Bearer <access-token>"
```

## 8. Deploy to Vercel

```bash
# Install Vercel CLI
npm i -g vercel

# Login
vercel login

# Deploy
vercel --prod
```

## 9. Создать Webhook Subscription

После deploy:

```bash
curl -X POST https://your-app.vercel.app/api/subscribe \
  -H "Content-Type: application/json" \
  -d '{
    "siteId": "kzprimeestate.sharepoint.com,9f482633-8093-471a-a0e4-c6353a265373,df166f91-e314-4b9b-87a7-cbd2e26f2c1c",
    "fileId": "<your-file-id>"
  }'
```

## 10. Проверка

1. Открой Excel в SharePoint
2. Измени любую ячейку
3. Сохрани
4. Проверь React app - данные должны обновиться автоматически!

## Troubleshooting

### Webhook не работает
- Проверь что URL публичный (HTTPS)
- Проверь Admin Consent в Azure
- Проверь logs в Vercel Dashboard

### Realtime не работает
- Проверь что таблица `sharepoint_updates` включена в Replication
- Проверь Service Role Key
- Проверь Browser Console для ошибок

### Subscription истекает
- Vercel Cron автоматически продлевает каждые 50 минут
- Проверь Cron logs в Vercel Dashboard
