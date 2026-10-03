import type { Mail, Folder, Attachment } from './common';

const uid = (p = 'm'): string =>
  p + '_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-3);

export type SendResult =
  | { ok: true }
  | { ok: false; status?: number; code?: string; message: string };

export async function updateRealMail(
  provider: 'gmail' | 'outlook',
  accessToken: string,
  messageId: string,
  action: 'star' | 'unstar' | 'trash' | 'spam' | 'restore' | 'delete'
): Promise<boolean> {
  try {
    const gmailUrl = `https://gmail.googleapis.com/gmail/v1/users/me/messages/${encodeURIComponent(messageId)}`;
    const outlookUrl = `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}`;
    const url = provider === 'gmail'
      ? action === 'delete' ? gmailUrl : `${gmailUrl}/modify`
      : action === 'delete' ? outlookUrl
        : action === 'trash' || action === 'spam' || action === 'restore' ? `${outlookUrl}/move`
          : outlookUrl;
    const body = provider === 'gmail'
      ? action === 'star' ? { addLabelIds: ['STARRED'] }
        : action === 'unstar' ? { removeLabelIds: ['STARRED'] }
          : action === 'trash' ? { addLabelIds: ['TRASH'], removeLabelIds: ['INBOX'] }
            : action === 'spam' ? { addLabelIds: ['SPAM'], removeLabelIds: ['INBOX'] }
              : action === 'restore' ? { addLabelIds: ['INBOX'], removeLabelIds: ['TRASH', 'SPAM'] }
                : undefined
      : action === 'trash' ? { destinationId: 'deleteditems' }
        : action === 'spam' ? { destinationId: 'junkemail' }
          : action === 'restore' ? { destinationId: 'inbox' }
            : action === 'star' || action === 'unstar' ? { flag: { flagStatus: action === 'star' ? 'flagged' : 'notFlagged' } }
              : undefined;
    const res = await fetch(url, {
      method: provider === 'gmail' ? action === 'delete' ? 'DELETE' : 'POST'
        : action === 'delete' ? 'DELETE' : action === 'trash' || action === 'spam' || action === 'restore' ? 'POST' : 'PATCH',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        Prefer: 'IdType="ImmutableId"',
      },
      ...(body ? { body: JSON.stringify(body) } : {}),
    });
    return res.ok;
  } catch {
    return false;
  }
}

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

function dataUrlParts(dataUrl: string): { type: string; base64: string } | null {
  const match = dataUrl.match(/^data:([^;,]+)?;base64,(.+)$/);
  return match ? { type: match[1] || 'application/octet-stream', base64: match[2] } : null;
}

function bodyBase64Placeholder(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

function wrapBase64(value: string): string {
  return value.replace(/.{1,76}/g, '$&\r\n').trimEnd();
}

function encodeMimeHeader(value: string): string {
  return /^[\x20-\x7E]*$/.test(value) ? value : encodeMimeSubject(value);
}

function decodeBase64Utf8(data: string): string {
  const normalized = data.replace(/-/g, '+').replace(/_/g, '/');
  const binary = atob(normalized + '='.repeat((4 - (normalized.length % 4)) % 4));
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  return new TextDecoder('utf-8', { fatal: false }).decode(bytes);
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
      const decoded = decodeBase64Utf8(payload.body.data);
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

function collectGmailAttachments(payload: any, result: any[] = []): any[] {
  if (!payload) return result;
  if (payload.filename && (payload.body?.attachmentId || payload.body?.data)) {
    result.push({
      name: payload.filename,
      size: Number(payload.body.size) || 0,
      type: payload.mimeType || 'application/octet-stream',
      attachmentId: payload.body.attachmentId,
      data: payload.body.data,
    });
  }
  for (const part of payload.parts || []) collectGmailAttachments(part, result);
  return result;
}

async function loadGmailAttachments(
  accessToken: string,
  messageId: string,
  payload: any,
): Promise<Attachment[]> {
  const attachments = collectGmailAttachments(payload);
  const loaded: Attachment[] = [];
  for (const attachment of attachments) {
    try {
      let base64 = String(attachment.data || '');
      if (!base64 && attachment.attachmentId) {
        const response = await fetch(
          `https://gmail.googleapis.com/gmail/v1/users/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachment.attachmentId)}`,
          { headers: { Authorization: `Bearer ${accessToken}` } },
        );
        if (!response.ok) continue;
        const data = await response.json();
        base64 = String(data.data || '');
      }
      base64 = base64.replace(/-/g, '+').replace(/_/g, '/');
      if (base64) {
        loaded.push({
          name: attachment.name,
          size: attachment.size,
          type: attachment.type,
          data: `data:${attachment.type};base64,${base64}`,
        });
      }
    } catch {
      // Keep the message available when one attachment cannot be downloaded.
    }
  }
  return loaded;
}

/**
 * Send email in real-time via Google Gmail API
 */
export async function sendRealGmailMail(
  accessToken: string,
  to: string,
  subject: string,
  bodyHtmlOrText: string,
  attachments: Attachment[] = []
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
    const usableAttachments = attachments
      .map((attachment) => ({ attachment, parts: attachment.data ? dataUrlParts(attachment.data) : null }))
      .filter((item): item is { attachment: Attachment; parts: { type: string; base64: string } } => !!item.parts);
    const boundary = `triade_${Date.now().toString(36)}`;
    const mimeBody = [
      `--${boundary}`,
      'Content-Type: text/html; charset=UTF-8',
      'Content-Transfer-Encoding: base64',
      '',
      wrapBase64(bodyBase64Placeholder(safeBody)),
      ...usableAttachments.flatMap(({ attachment, parts }) => [
        `--${boundary}`,
        `Content-Type: ${parts.type}; name="${encodeMimeHeader(attachment.name)}"`,
        `Content-Disposition: attachment; filename="${encodeMimeHeader(attachment.name)}"`,
        'Content-Transfer-Encoding: base64',
        '',
        wrapBase64(parts.base64),
      ]),
      `--${boundary}--`,
    ].join('\r\n');

    const rawMessage = [
      `To: ${to}`,
      `Subject: ${encodeMimeSubject(subject || '(Sin asunto)')}`,
      'MIME-Version: 1.0',
      `Content-Type: multipart/mixed; boundary="${boundary}"`,
      '',
      mimeBody,
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
  bodyHtmlOrText: string,
  attachments: Attachment[] = []
): Promise<SendResult> {
  try {
    if (!accessToken || !accessToken.trim()) {
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
          attachments: attachments
            .filter((attachment) => attachment.data && dataUrlParts(attachment.data))
            .map((attachment) => {
              const parts = dataUrlParts(attachment.data!);
              return {
                '@odata.type': '#microsoft.graph.fileAttachment',
                name: attachment.name,
                contentType: parts!.type,
                contentBytes: parts!.base64,
              };
            }),
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
export async function fetchRealGmailMails(
  accessToken: string,
  accountEmail: string,
  onBatch?: (batch: Mail[]) => void,
): Promise<Mail[]> {
  try {
    const messageSummaries: { id: string; threadId: string }[] = [];
    let pageToken = '';
    for (let page = 0; page < 20; page += 1) {
      const params = new URLSearchParams({ maxResults: '100', includeSpamTrash: 'true' });
      if (pageToken) params.set('pageToken', pageToken);
      const listRes = await fetch(`https://gmail.googleapis.com/gmail/v1/users/me/messages?${params}`, {
        headers: { Authorization: `Bearer ${accessToken}` },
      });
      if (!listRes.ok) {
        const errText = await listRes.text();
        console.warn('Error fetching Gmail messages list:', listRes.status, errText);
        throw new Error(`Gmail API HTTP ${listRes.status}: ${errText}`);
      }
      const listData = await listRes.json();
      messageSummaries.push(...(listData.messages || []));
      pageToken = listData.nextPageToken || '';
      if (!pageToken) break;
    }

    if (!messageSummaries.length) return [];

    const fetchedMails: Mail[] = [];

    // Fetch details in small concurrent batches so the UI can render progressively.
    const batchSize = 8;
    for (let offset = 0; offset < messageSummaries.length; offset += batchSize) {
      const summaries = messageSummaries.slice(offset, offset + batchSize);
      const batch = await Promise.all(summaries.map(async (item): Promise<Mail | null> => {
        try {
        const detailRes = await fetch(
          `https://gmail.googleapis.com/gmail/v1/users/me/messages/${item.id}?format=full`,
          {
            headers: { Authorization: `Bearer ${accessToken}` },
          }
        );
        if (!detailRes.ok) return null;

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
        if (labelIds.includes('TRASH')) {
          folder = 'trash';
        } else if (labelIds.includes('SPAM')) {
          folder = 'spam';
        } else if (labelIds.includes('DRAFT')) {
          folder = 'drafts';
        } else if (labelIds.includes('SENT')) {
          folder = 'sent';
        }

        const { text: bodyText, html: bodyHtml } = extractGmailBody(msg.payload);
        const attachments = await loadGmailAttachments(accessToken, item.id, msg.payload);
        const finalBodyText = bodyText || snippet || '(Sin contenido)';
        const finalBodyHtml = bodyHtml || undefined;

        return {
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
          attachments,
        };
        } catch (err) {
          console.warn(`Failed to parse Gmail message ${item.id}:`, err);
          return null;
        }
      }));
      const successfulBatch = batch.filter((mail): mail is Mail => mail !== null);
      fetchedMails.push(...successfulBatch);
      if (successfulBatch.length) onBatch?.(successfulBatch);
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
async function loadOutlookAttachment(accessToken: string, messageId: string, attachment: any): Promise<Attachment | null> {
  if (!attachment || attachment.isInline || attachment.contentType === 'message/rfc822') {
    return null;
  }

  if (attachment.contentBytes) {
    return {
      name: attachment.name || 'Adjunto',
      size: Number(attachment.size) || 0,
      type: attachment.contentType || 'application/octet-stream',
      data: `data:${attachment.contentType || 'application/octet-stream'};base64,${attachment.contentBytes}`,
    };
  }

  if (!attachment.id) {
    return null;
  }

  try {
    const attachmentRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachment.id)}`,
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      }
    );

    if (!attachmentRes.ok) {
      return null;
    }

    const attachmentData = await attachmentRes.json();
    if (!attachmentData.contentBytes) {
      return null;
    }

    return {
      name: attachmentData.name || attachment.name || 'Adjunto',
      size: Number(attachmentData.size) || Number(attachment.size) || 0,
      type: attachmentData.contentType || attachment.contentType || 'application/octet-stream',
      data: `data:${attachmentData.contentType || attachment.contentType || 'application/octet-stream'};base64,${attachmentData.contentBytes}`,
    };
  } catch (err) {
    console.warn('Failed to fetch Outlook attachment bytes:', err);
    return null;
  }
}

async function loadOutlookMessageAttachments(accessToken: string, messageId: string): Promise<Attachment[]> {
  try {
    const listRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/attachments?$select=id,name,contentType,size,isInline`,
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      }
    );

    if (!listRes.ok) {
      return [];
    }

    const listData = await listRes.json();
    const attachments: any[] = listData.value || [];
    const loaded: Attachment[] = [];

    for (const attachment of attachments) {
      if (!attachment || attachment.isInline) continue;
      const resolved = await loadOutlookAttachment(accessToken, messageId, attachment);
      if (resolved) {
        loaded.push(resolved);
      }
    }

    return loaded;
  } catch (err) {
    console.warn('Failed to fetch Outlook message attachments:', err);
    return [];
  }
}

export async function fetchRealOutlookMails(accessToken: string, accountEmail: string): Promise<Mail[]> {
  try {
    const folders: { id: string; folder: Folder }[] = [
      { id: 'inbox', folder: 'inbox' },
      { id: 'sentitems', folder: 'sent' },
      { id: 'drafts', folder: 'drafts' },
      { id: 'junkemail', folder: 'spam' },
      { id: 'deleteditems', folder: 'trash' },
    ];
    const items: { message: any; folder: Folder }[] = [];
    const select = 'id,subject,bodyPreview,body,from,toRecipients,receivedDateTime,sentDateTime,isRead,flag,isDraft';
    for (const folder of folders) {
      let nextUrl: string | null = `https://graph.microsoft.com/v1.0/me/mailFolders/${folder.id}/messages?$top=100&$select=${select}`;
      for (let page = 0; nextUrl && page < 10; page += 1) {
        const pageUrl: string = nextUrl;
        const res: Response = await fetch(pageUrl, {
          headers: { Authorization: `Bearer ${accessToken}`, Prefer: 'IdType="ImmutableId"' },
        });
        if (!res.ok) {
          const errText = await res.text();
          console.warn(`Error fetching Outlook ${folder.id} messages:`, res.status, errText);
          throw new Error(`Microsoft Graph API HTTP ${res.status}: ${errText}`);
        }
        const data: { value?: any[]; '@odata.nextLink'?: string } = await res.json();
        items.push(...(data.value || []).map((message: any) => ({ message, folder: folder.folder })));
        nextUrl = data['@odata.nextLink'] || null;
      }
    }

    const fetchedMails: Mail[] = [];

    for (const { message: m, folder } of items) {
      const senderName = m.from?.emailAddress?.name || m.from?.emailAddress?.address || 'Outlook User';
      const senderEmail = m.from?.emailAddress?.address || '';
      const toAddress = m.toRecipients?.[0]?.emailAddress?.address || accountEmail;
      const dateObj = m.receivedDateTime ? new Date(m.receivedDateTime) : (m.sentDateTime ? new Date(m.sentDateTime) : new Date());
      const timeLabel = isNaN(dateObj.getTime())
        ? new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
        : dateObj.toLocaleDateString([], { month: 'short', day: 'numeric' }) + ' ' + dateObj.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

      const attachments = await loadOutlookMessageAttachments(accessToken, m.id || '');

      const timestampNum = !isNaN(dateObj.getTime()) ? dateObj.getTime() : Date.now();

      fetchedMails.push({
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
        folder: m.isDraft ? 'drafts' : folder,
        attachments,
      });
    }

    // Ordenar de más reciente a más antiguo (timestamp descendente)
    return fetchedMails.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
  } catch (err) {
    console.error('fetchRealOutlookMails error:', err);
    throw err;
  }
}
