import { ConfidentialClientApplication } from '@azure/msal-node';
import { Client } from '@microsoft/microsoft-graph-client';
import { createClient } from '@supabase/supabase-js';
import 'isomorphic-fetch';

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY
);

const msalConfig = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
    clientSecret: process.env.AZURE_CLIENT_SECRET
  }
};

async function getGraphClient() {
  const cca = new ConfidentialClientApplication(msalConfig);
  const result = await cca.acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default']
  });

  return Client.init({
    authProvider: (done) => {
      done(null, result.accessToken);
    }
  });
}

export default async function handler(req, res) {
  try {
    // Получаем активную подписку из Supabase
    const { data: subscriptions } = await supabase
      .from('graph_subscriptions')
      .select('*')
      .eq('is_active', true)
      .single();

    if (!subscriptions) {
      return res.status(200).json({ message: 'No active subscriptions' });
    }

    const graphClient = await getGraphClient();

    // Продлеваем подписку
    const newExpiration = new Date(Date.now() + 3600000).toISOString();

    await graphClient
      .api(`/subscriptions/${subscriptions.subscription_id}`)
      .patch({
        expirationDateTime: newExpiration
      });

    // Обновляем в Supabase
    await supabase
      .from('graph_subscriptions')
      .update({ expires_at: newExpiration })
      .eq('id', subscriptions.id);

    console.log('🔄 Subscription renewed');

    res.status(200).json({ success: true, expiresAt: newExpiration });
  } catch (error) {
    console.error('❌ Renew error:', error);
    res.status(500).json({ error: error.message });
  }
}
