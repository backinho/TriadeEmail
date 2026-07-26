import type { Mail, Folder, EmailProvider } from './common';

const uid = (p = 'm'): string =>
  p + '_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-3);

/**
 * Extracts real live emails from Google Gmail API
 * Endpoint: https://gmail.googleapis.com/gmail/v1/users/me/messages
 */
/**
 * Extracts real live emails from Google Gmail API
 * Endpoint: https://gmail.googleapis.com/gmail/v1/users/me/messages
 */
export async function fetchRealGmailMails(accessToken: string, accountEmail: string): Promise<Mail[]> {
  try {
    const listRes = await fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/messages?maxResults=20',
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      }
    );

    if (!listRes.ok) {
      const errText = await listRes.text();
      console.warn('Error fetching Gmail messages list:', listRes.status, errText);
      throw new Error(`Gmail API HTTP ${listRes.status}: ${errText}`);
    }

    const listData = await listRes.json();
    const messageSummaries: { id: string; threadId: string }[] = listData.messages || [];

    if (!messageSummaries.length) return [];

    const fetchedMails: Mail[] = [];

    // Fetch detail for top 20 messages
    for (const item of messageSummaries.slice(0, 20)) {
      try {
        const detailRes = await fetch(
          `https://gmail.googleapis.com/gmail/v1/users/me/messages/${item.id}?format=full`,
          {
            headers: { Authorization: `Bearer ${accessToken}` },
          }
        );
        if (!detailRes.ok) continue;

        const msg = await detailRes.json();
        const headers: { name: string; value: string }[] = msg.payload?.headers || [];

        const subjectHeader = headers.find((h) => h.name.toLowerCase() === 'subject');
        const fromHeader = headers.find((h) => h.name.toLowerCase() === 'from');
        const dateHeader = headers.find((h) => h.name.toLowerCase() === 'date');

        const rawFrom = fromHeader?.value || 'Remitente';
        const subject = subjectHeader?.value || '(Sin asunto)';
        const snippet = msg.snippet || '';

        let senderName = rawFrom;
        let senderEmail = rawFrom;
        const match = rawFrom.match(/^(?:"?([^"]*)"?\s)?<([^>]+)>$/);
        if (match) {
          senderName = match[1] || match[2];
          senderEmail = match[2];
        }

        const dateObj = dateHeader ? new Date(dateHeader.value) : new Date();
        const timeLabel = isNaN(dateObj.getTime())
          ? new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
          : dateObj.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

        const isUnread = Boolean(msg.labelIds?.includes('UNREAD'));
        const isStarred = Boolean(msg.labelIds?.includes('STARRED'));

        let bodyContent = snippet;
        if (msg.payload?.body?.data) {
          try {
            bodyContent = atob(msg.payload.body.data.replace(/-/g, '+').replace(/_/g, '/'));
          } catch {
            bodyContent = snippet;
          }
        }

        fetchedMails.push({
          id: item.id || uid('m'),
          from: senderName,
          fromEmail: senderEmail,
          to: accountEmail,
          subject: subject,
          body: bodyContent || snippet || '(Sin contenido)',
          account: accountEmail,
          time: timeLabel,
          unread: isUnread,
          starred: isStarred,
          folder: 'inbox',
        });
      } catch (err) {
        console.warn(`Failed to parse Gmail message ${item.id}:`, err);
      }
    }

    return fetchedMails;
  } catch (err) {
    console.error('fetchRealGmailMails error:', err);
    throw err;
  }
}

/**
 * Extracts real live emails from Microsoft Graph API (Outlook)
 * Endpoint: https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages
 */
export async function fetchRealOutlookMails(accessToken: string, accountEmail: string): Promise<Mail[]> {
  try {
    const res = await fetch(
      'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=20&$select=id,subject,bodyPreview,body,from,receivedDateTime,isRead,flag',
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      }
    );

    if (!res.ok) {
      const errText = await res.text();
      console.warn('Error fetching Outlook messages:', res.status, errText);
      throw new Error(`Microsoft Graph API HTTP ${res.status}: ${errText}`);
    }

    const data = await res.json();
    const items: any[] = data.value || [];

    const fetchedMails: Mail[] = items.map((m: any) => {
      const senderName = m.from?.emailAddress?.name || m.from?.emailAddress?.address || 'Outlook User';
      const senderEmail = m.from?.emailAddress?.address || '';
      const dateObj = m.receivedDateTime ? new Date(m.receivedDateTime) : new Date();
      const timeLabel = isNaN(dateObj.getTime())
        ? new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
        : dateObj.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

      return {
        id: uid('m'),
        from: senderName,
        fromEmail: senderEmail,
        to: accountEmail,
        subject: m.subject || '(Sin asunto)',
        body: m.bodyPreview || (m.body?.content ? m.body.content.replace(/<[^>]+>/g, '').slice(0, 200) : ''),
        account: accountEmail,
        time: timeLabel,
        unread: !m.isRead,
        starred: m.flag?.flagStatus === 'flagged',
        folder: 'inbox' as Folder,
      };
    });

    return fetchedMails;
  } catch (err) {
    console.error('fetchRealOutlookMails error:', err);
    throw err;
  }
}
