-- Таблица для хранения обновлений из SharePoint
CREATE TABLE IF NOT EXISTS sharepoint_updates (
  id BIGSERIAL PRIMARY KEY,
  resource TEXT NOT NULL,
  change_type TEXT NOT NULL,
  subscription_id TEXT,
  client_state TEXT,
  notification_data JSONB,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Таблица для хранения активных подписок
CREATE TABLE IF NOT EXISTS graph_subscriptions (
  id BIGSERIAL PRIMARY KEY,
  subscription_id TEXT UNIQUE NOT NULL,
  resource TEXT NOT NULL,
  expires_at TIMESTAMP WITH TIME ZONE NOT NULL,
  is_active BOOLEAN DEFAULT true,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Включаем Realtime для sharepoint_updates
ALTER PUBLICATION supabase_realtime ADD TABLE sharepoint_updates;

-- Индексы для производительности
CREATE INDEX idx_sharepoint_updates_created_at ON sharepoint_updates(created_at DESC);
CREATE INDEX idx_graph_subscriptions_active ON graph_subscriptions(is_active) WHERE is_active = true;
