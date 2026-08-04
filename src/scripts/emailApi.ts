import type { Mail, Folder } from './common';

const uid = (p = 'm'): string =>
  p + '_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-3);

export type SendResult =
  | { ok: true }
  | { ok: false; status?: number; code?: string; message: string };

function toBase64Url(str: string): string {
  const bytes = new TextEncoder().encode(str);
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function encodeMimeSubject(subject: string): string {
  if (/^[\x20-\x7E]*$/.test(subject)) return subject;
  const bytes = new TextEncoder().encode(subject);
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return `=?UTF-8?B?${btoa(binary)}?=`;
}

function parseApiError(status: number, errText: string): string {
  try {
    const data = JSON.parse(errText);
    const msg = data?.error?.message || data?.error_description || data?.message;
    if (msg) return String(msg);
  } catch {
    // ignore
  }
  return errText.slice(0, 200) || `HTTP ${status}`;
}

export async function checkGoogleTokenScopes(accessToken: string): Promise<{ valid: boolean; hasSend: boolean; scopes: string }> {
  try {
    const res = await fetch(`https://www.googleapis.com/oauth2/v1/tokeninfo?access_token=${encodeURIComponent(accessToken)}`);
    if (!res.ok) return { valid: false, hasSend: false, scopes: '' };
    const data = await res.json();
    const scopes = String(data.scope || '');
    return {
      valid: true,
      hasSend: scopes.includes('gmail.send') || scopes.includes('mail.google.com') || scopes.includes('gmail.compose'),
      scopes,
    };
  } catch {
    return { valid: false, hasSend: false, scopes: '' };
  }
}

/**
 * Recurse through Gmail payload parts to extract body text and html
 */
function extractGmailBody(payload: any): { text: string; html: string } {
  let text = '';
  let html = '';

  if (!payload) return { text, html };

  if (payload.body?.data) {
    try {
      const decoded = atob(payload.body.data.replace(/-/g, '+').replace(/_/g, '/'));
      if (payload.mimeType === 'text/html') {
        html = decoded;
      } else {
        text = decoded;
      }
    } catch {
      // Ignore decoding errors
    }
  }

  if (payload.parts && Array.isArray(payload.parts)) {
    for (const part of payload.parts) {
      const res = extractGmailBody(part);
      if (res.html && !html) html = res.html;
      if (res.text && !text) text = res.text;
    }
  }

  return { text, html };
}

/**
 * Send email in real-time via Google Gmail API
 */
export async function sendRealGmailMail(
  accessToken: string,
  to: string,
  subject: string,
  bodyHtmlOrText: string
): Promise<SendResult> {
  try {
    if (!accessToken.startsWith('ya29')) {
      return { ok: false, code: 'invalid_token', message: 'Token de Google inválido. Vuelve a vincular la cuenta.' };
    }

    const scopeCheck = await checkGoogleTokenScopes(accessToken);
    if (scopeCheck.valid && !scopeCheck.hasSend) {
      return {
        ok: false,
        code: 'missing_scope',
        message: 'La cuenta no tiene permiso de envío (gmail.send). Desvincúlala y vuelve a vincularla.',
      };
    }

    const safeBody = bodyHtmlOrText || '<p></p>';
    const bodyBytes = new TextEncoder().encode(safeBody);
    let bodyBinary = '';
    for (const byte of bodyBytes) bodyBinary += String.fromCharCode(byte);
    const bodyBase64 = btoa(bodyBinary);

    const rawMessage = [
      `To: ${to}`,
      `Subject: ${encodeMimeSubject(subject || '(Sin asunto)')}`,
      'MIME-Version: 1.0',
      'Content-Type: text/html; charset=UTF-8',
      'Content-Transfer-Encoding: base64',
      '',
      bodyBase64,
    ].join('\r\n');

    const res = await fetch('https://gmail.googleapis.com/gmail/v1/users/me/messages/send', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ raw: toBase64Url(rawMessage) }),
    });

    if (!res.ok) {
      const errText = await res.text();
      console.warn('Gmail API Send error:', res.status, errText);
      const detail = parseApiError(res.status, errText);
      if (res.status === 403) {
        return { ok: false, status: 403, code: 'missing_scope', message: `Permiso denegado: ${detail}. Desvincula y vuelve a vincular la cuenta.` };
      }
      if (res.status === 401) {
        return { ok: false, status: 401, code: 'expired_token', message: 'Token expirado. Vuelve a vincular la cuenta de Gmail.' };
      }
      return { ok: false, status: res.status, message: detail };
    }
    return { ok: true };
  } catch (err) {
    console.error('sendRealGmailMail exception:', err);
    return { ok: false, message: err instanceof Error ? err.message : 'Error desconocido al enviar' };
  }
}

/**
 * Send email in real-time via Microsoft Graph API
 */
export async function sendRealOutlookMail(
  accessToken: string,
  to: string,
  subject: string,
  bodyHtmlOrText: string
): Promise<SendResult> {
  try {
    if (accessToken.startsWith('eyJ')) {
      return { ok: false, code: 'invalid_token', message: 'Token de Microsoft inválido. Vuelve a vincular la cuenta.' };
    }

    const res = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        message: {
          subject: subject || '(Sin asunto)',
          body: {
            contentType: 'HTML',
            content: bodyHtmlOrText || '<p></p>',
          },
          toRecipients: [
            {
              emailAddress: { address: to },
            },
          ],
        },
        saveToSentItems: true,
      }),
    });

    if (!res.ok) {
      const errText = await res.text();
      console.warn('Outlook API Send error:', res.status, errText);
      const detail = parseApiError(res.status, errText);
      if (res.status === 403) {
        return { ok: false, status: 403, code: 'missing_scope', message: `Permiso denegado: ${detail}. Desvincula y vuelve a vincular la cuenta.` };
      }
      if (res.status === 401) {
        return { ok: false, status: 401, code: 'expired_token', message: 'Token expirado. Vuelve a vincular la cuenta de Outlook.' };
      }
      return { ok: false, status: res.status, message: detail };
    }
    return { ok: true };
  } catch (err) {
    console.error('sendRealOutlookMail exception:', err);
    return { ok: false, message: err instanceof Error ? err.message : 'Error desconocido al enviar' };
  }
}

/**
 * Extracts real live emails from Google Gmail API (Inbox, Sent, Drafts, Spam, Trash)
 * Endpoint: https://gmail.googleapis.com/gmail/v1/users/me/messages
 */
export async function fetchRealGmailMails(accessToken: string, accountEmail: string): Promise<Mail[]> {
  try {
    const listRes = await fetch(
      'https://gmail.googleapis.com/gmail/v1/users/me/messages?maxResults=50',
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

    // Fetch detail for messages
    for (const item of messageSummaries.slice(0, 50)) {
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
        const toHeader = headers.find((h) => h.name.toLowerCase() === 'to');
        const dateHeader = headers.find((h) => h.name.toLowerCase() === 'date');

        const rawFrom = fromHeader?.value || 'Remitente';
        const rawTo = toHeader?.value || accountEmail;
        const subject = subjectHeader?.value || '(Sin asunto)';
        const snippet = msg.snippet || '';

        let senderName = rawFrom;
        let senderEmail = rawFrom;
        const matchFrom = rawFrom.match(/^(?:"?([^"]*)"?\s)?<([^>]+)>$/);
        if (matchFrom) {
          senderName = matchFrom[1] || matchFrom[2];
          senderEmail = matchFrom[2];
        }

        let toEmail = rawTo;
        const matchTo = rawTo.match(/^(?:"?([^"]*)"?\s)?<([^>]+)>$/);
        if (matchTo) {
          toEmail = matchTo[2];
        }

        const internalTs = Number(msg.internalDate) || (dateHeader ? new Date(dateHeader.value).getTime() : Date.now());
        const dateObj = new Date(internalTs);
        const isToday = dateObj.toDateString() === new Date().toDateString();
        const timeLabel = isNaN(dateObj.getTime())
          ? new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
          : isToday
            ? dateObj.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
            : dateObj.toLocaleDateString([], { month: 'short', day: 'numeric' }) + ', ' + dateObj.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

        const labelIds: string[] = msg.labelIds || [];
        const isUnread = labelIds.includes('UNREAD');
        const isStarred = labelIds.includes('STARRED');

        let folder: Folder = 'inbox';
        if (labelIds.includes('SENT')) {
          folder = 'sent';
        } else if (labelIds.includes('DRAFT')) {
          folder = 'drafts';
        } else if (labelIds.includes('SPAM')) {
          folder = 'spam';
        } else if (labelIds.includes('TRASH')) {
          folder = 'trash';
        }

        const { text: bodyText, html: bodyHtml } = extractGmailBody(msg.payload);
        const finalBodyText = bodyText || snippet || '(Sin contenido)';
        const finalBodyHtml = bodyHtml || undefined;

        fetchedMails.push({
          id: item.id || uid('m'),
          from: senderName,
          fromEmail: senderEmail,
          to: toEmail,
          subject: subject,
          body: finalBodyText,
          bodyHtml: finalBodyHtml,
          account: accountEmail,
          time: timeLabel,
          timestamp: internalTs,
          unread: isUnread,
          starred: isStarred,
          folder: folder,
        });
      } catch (err) {
        console.warn(`Failed to parse Gmail message ${item.id}:`, err);
      }
    }

    // Ordenar siempre de más reciente a más antiguo (timestamp descendente)
    return fetchedMails.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
  } catch (err) {
    console.error('fetchRealGmailMails error:', err);
    throw err;
  }
}

/**
 * Extracts real live emails from Microsoft Graph API (Outlook)
 * Endpoint: https://graph.microsoft.com/v1.0/me/messages
 */
export async function fetchRealOutlookMails(accessToken: string, accountEmail: string): Promise<Mail[]> {
  try {
    const res = await fetch(
      'https://graph.microsoft.com/v1.0/me/messages?$top=50&$select=id,subject,bodyPreview,body,from,toRecipients,receivedDateTime,sentDateTime,isRead,flag',
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
      const toAddress = m.toRecipients?.[0]?.emailAddress?.address || accountEmail;
      const dateObj = m.receivedDateTime ? new Date(m.receivedDateTime) : (m.sentDateTime ? new Date(m.sentDateTime) : new Date());
      const timeLabel = isNaN(dateObj.getTime())
        ? new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
        : dateObj.toLocaleDateString([], { month: 'short', day: 'numeric' }) + ' ' + dateObj.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

      const isSent = senderEmail.toLowerCase() === accountEmail.toLowerCase();

      const timestampNum = !isNaN(dateObj.getTime()) ? dateObj.getTime() : Date.now();

      return {
        id: m.id || uid('m'),
        from: senderName,
        fromEmail: senderEmail,
        to: toAddress,
        subject: m.subject || '(Sin asunto)',
        body: m.bodyPreview || (m.body?.content ? m.body.content.replace(/<[^>]+>/g, '').slice(0, 200) : ''),
        bodyHtml: m.body?.contentType === 'html' ? m.body.content : undefined,
        account: accountEmail,
        time: timeLabel,
        timestamp: timestampNum,
        unread: !m.isRead,
        starred: m.flag?.flagStatus === 'flagged',
        folder: isSent ? ('sent' as Folder) : ('inbox' as Folder),
      };
    });

    // Ordenar de más reciente a más antiguo (timestamp descendente)
    return fetchedMails.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
  } catch (err) {
    console.error('fetchRealOutlookMails error:', err);
    throw err;
  }
}
