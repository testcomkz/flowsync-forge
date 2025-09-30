import { ConfidentialClientApplication } from '@azure/msal-node';
import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';

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
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { siteId, fileId } = req.body;

    if (!siteId || !fileId) {
      return res.status(400).json({ error: 'siteId and fileId required' });
    }

    const graphClient = await getGraphClient();

    const subscription = {
      changeType: 'updated',
      notificationUrl: `${process.env.VERCEL_URL || process.env.WEBHOOK_URL}/api/webhook`,
      resource: `/sites/${siteId}/drive/items/${fileId}`,
      expirationDateTime: new Date(Date.now() + 3600000).toISOString(), // 1 час
      clientState: process.env.WEBHOOK_SECRET
    };

    const result = await graphClient
      .api('/subscriptions')
      .post(subscription);

    console.log('✅ Subscription created:', result.id);

    res.status(200).json({
      success: true,
      subscriptionId: result.id,
      expiresAt: result.expirationDateTime
    });
  } catch (error) {
    console.error('❌ Subscribe error:', error);
    res.status(500).json({ error: error.message });
  }
}
