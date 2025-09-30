import { createClient } from '@supabase/supabase-js';

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY
);

export default async function handler(req, res) {
  // Validation token для регистрации webhook
  if (req.query.validationToken) {
    console.log('📨 Webhook validation');
    return res.status(200).send(req.query.validationToken);
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const notifications = req.body.value;

    if (!notifications || !Array.isArray(notifications)) {
      return res.status(400).json({ error: 'Invalid notification' });
    }

    console.log(`📬 Received ${notifications.length} notification(s)`);

    for (const notification of notifications) {
      // Проверка clientState
      if (notification.clientState !== process.env.WEBHOOK_SECRET) {
        console.warn('⚠️ Invalid clientState');
        continue;
      }

      // Сохраняем в Supabase
      const { error: insertError } = await supabase
        .from('sharepoint_updates')
        .insert({
          resource: notification.resource,
          change_type: notification.changeType,
          subscription_id: notification.subscriptionId,
          client_state: notification.clientState,
          notification_data: notification
        });

      if (insertError) {
        console.error('❌ Supabase insert error:', insertError);
      } else {
        console.log('✅ Update saved to Supabase');
      }

      // Отправляем broadcast через Realtime
      const channel = supabase.channel('sharepoint-updates');
      await channel.send({
        type: 'broadcast',
        event: 'excel-updated',
        payload: {
          resource: notification.resource,
          changeType: notification.changeType,
          timestamp: new Date().toISOString()
        }
      });
      console.log('📡 Broadcast sent to all clients');
    }

    res.status(202).json({ status: 'accepted' });
  } catch (error) {
    console.error('❌ Webhook error:', error);
    res.status(500).json({ error: error.message });
  }
}
