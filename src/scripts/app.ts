import {
  KEYS, store, t, applyTheme, applyI18n,
  type Account, type Category, type Profile, type Mail, type Attachment, type Folder, type EmailProvider,
} from './common';
import { supabase, supabaseUrl, supabaseAnonKey } from './supabase';
import { fetchRealGmailMails, fetchRealOutlookMails, sendRealGmailMail, sendRealOutlookMail, type SendResult } from './emailApi';


// ---------- Defaults ----------
const DEFAULT_ACCOUNTS: Account[] = [];
const DEFAULT_CATEGORIES: Category[] = [];
const DEFAULT_PROFILE: Profile = {
  name: 'Usuario Triade',
  email: '',
  phone: '',
  signature: '— Enviado desde Triade Mail',
};

const GOOGLE_OAUTH_SCOPES = [
  'https://www.googleapis.com/auth/gmail.readonly',
  'https://www.googleapis.com/auth/gmail.send',
  'https://www.googleapis.com/auth/userinfo.email',
].join(' ');

const MICROSOFT_OAUTH_SCOPES = [
  'https://graph.microsoft.com/Mail.Read',
  'https://graph.microsoft.com/Mail.Send',
  'https://graph.microsoft.com/User.Read',
].join(' ');

const uid = (p = 'm'): string =>
  p + '_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-3);

// ---------- State ----------
function loadSessionAccounts(): Account[] {
  try {
    const raw = sessionStorage.getItem(KEYS.sessionAccounts);
    if (!raw) return [];
    const parsed = JSON.parse(raw) as Account[];
    return parsed.filter((a) => a.provider === 'gmail' || a.provider === 'outlook');
  } catch {
    return [];
  }
}

function saveSessionAccounts(accs: Account[]): void {
  try {
    const sanitized = accs.map(({ id, user_id, email, primary, provider, status, lastSync, access_token, refresh_token, expires_at, token }) => ({
      id, user_id, email, primary, provider, status, lastSync, access_token, refresh_token, expires_at, token,
    }));
    sessionStorage.setItem(KEYS.sessionAccounts, JSON.stringify(sanitized));
  } catch (err) {
    console.warn('No se pudo guardar cuentas en sessionStorage:', err);
  }
}

let accounts: Account[] = loadSessionAccounts();
if (accounts.length === 0) {
  accounts = store.get<Account[]>(KEYS.accounts, []).filter((a) => a.provider === 'gmail' || a.provider === 'outlook');
}
let categories: Category[] = store.get<Category[]>(KEYS.categories, DEFAULT_CATEGORIES);
let profile: Profile = store.get<Profile>(KEYS.profile, DEFAULT_PROFILE);
let pageSize: number = store.get<number>(KEYS.pageSize, 10);
let mails: Mail[] = [];

if (accounts.length > 0) {
  const validAccountEmails = new Set(accounts.map((a) => a.email));
  mails = store.get<Mail[]>(KEYS.mails, []).filter((m) => m.account && validAccountEmails.has(m.account));
} else {
  accounts = [];
  mails = [];
  store.del(KEYS.accounts);
  store.del(KEYS.mails);
}

let activeFolder: string = 'inbox';
let activeAccount: string = 'all';
let activeTab: string = 'all';
let searchQuery = '';
let currentPage = 1;

let currentUser: any = null;
let cachedAuthToken: string | null = null;
let composeFromAccount: string = '';

const persist = (): void => {
  saveSessionAccounts(accounts);
  store.set(KEYS.accounts, accounts.map(({ access_token, refresh_token, token, ...rest }) => ({
    ...rest,
    access_token: undefined,
    refresh_token: undefined,
    token: undefined,
  })));
  store.set(KEYS.categories, categories);
  store.set(KEYS.profile, profile);
  store.set(KEYS.pageSize, pageSize);

  // Guardar un resumen ligero de los 30 correos más recientes en localStorage para no exceder los 5MB del navegador
  const lightMails = mails.slice(0, 30).map((m) => ({
    id: m.id,
    from: m.from,
    fromEmail: m.fromEmail,
    to: m.to,
    subject: m.subject,
    body: (m.body || '').slice(0, 180),
    account: m.account,
    time: m.time,
    timestamp: m.timestamp,
    unread: m.unread,
    starred: m.starred,
    folder: m.folder,
  }));
  store.set(KEYS.mails, lightMails);

  if (currentUser) {
    supabase.from('profiles').upsert({
      id: currentUser.id,
      name: profile.name,
      email: profile.email,
      phone: profile.phone,
      signature: profile.signature,
      page_size: pageSize,
    }).then(({ error }) => {
      if (error) {
        // Fallback: Si RLS en Supabase no tiene política de INSERT habilitada, actualizar por id
        supabase.from('profiles').update({
          name: profile.name,
          email: profile.email,
          phone: profile.phone,
          signature: profile.signature,
          page_size: pageSize,
        }).eq('id', currentUser.id).then(({ error: updateErr }) => {
          if (updateErr) console.warn('Error al actualizar perfil en Supabase:', updateErr);
        });
      }
    });

    for (const cat of categories) {
      supabase.from('categories').upsert({
        user_id: currentUser.id,
        name: cat.name,
        color: cat.color,
        keywords: cat.keywords,
      }, { onConflict: 'user_id,name' }).then(({ error }) => {
        if (error) console.warn('Error al guardar categoría en Supabase:', error);
      });
    }
  }
};

function showPendingUnlinkNotices(): void {
  const notice = store.get<{ emails: string[]; at: number } | null>(KEYS.unlinkNotice, null);
  if (!notice?.emails?.length) return;
  store.del(KEYS.unlinkNotice);
  const emails = notice.emails.join(', ');
  if (notice.emails.length === 1) {
    toast(`ℹ️ ${t('account_unlinked_on_close').replace('{email}', notice.emails[0])}`);
  } else {
    toast(`ℹ️ ${t('accounts_unlinked_on_close').replace('{emails}', emails)}`);
  }
}

async function restoreAccountsToDb(): Promise<void> {
  if (!currentUser || accounts.length === 0) return;
  for (const acc of accounts) {
    const effToken = acc.access_token || acc.token;
    await supabase.from('user_accounts').upsert({
      user_id: currentUser.id,
      email: acc.email.toLowerCase(),
      access_token: effToken || null,
      refresh_token: acc.refresh_token || null,
      expires_at: acc.expires_at || null,
      is_primary: acc.primary || false,
    }, { onConflict: 'user_id,email' });
  }
}

function deleteAccountsFromDbKeepalive(userId: string, authToken: string, emails: string[]): void {
  for (const email of emails) {
    const url = `${supabaseUrl}/rest/v1/user_accounts?user_id=eq.${encodeURIComponent(userId)}&email=eq.${encodeURIComponent(email.toLowerCase())}`;
    fetch(url, {
      method: 'DELETE',
      headers: {
        apikey: supabaseAnonKey,
        Authorization: `Bearer ${authToken}`,
        Prefer: 'return=minimal',
      },
      keepalive: true,
    }).catch(() => {});
  }
}

async function unlinkAccountByEmail(
  removedEmail: string,
  options: { silent?: boolean; skipConfirm?: boolean; fromBrowserClose?: boolean } = {}
): Promise<boolean> {
  const emailLower = removedEmail.toLowerCase().trim();
  const index = accounts.findIndex((a) => a.email.toLowerCase() === emailLower);
  if (index === -1) return false;

  if (!options.skipConfirm && !options.fromBrowserClose) {
    const confirmed = await showConfirm({
      title: t('remove'),
      message: t('confirm_unlink').replace('{email}', emailLower),
      confirmText: t('remove'),
      danger: true,
    });
    if (!confirmed) return false;
  }

  const removed = accounts[index];
  accounts.splice(index, 1);

  if (removed.primary && accounts.length > 0) {
    accounts[0].primary = true;
    if (profile.email.toLowerCase() === emailLower) {
      profile.email = accounts[0].email;
    }
  }

  mails = mails.filter((m) => m.account.toLowerCase() !== emailLower);

  if (activeAccount.toLowerCase() === emailLower) {
    activeAccount = 'all';
  }

  const unlinked = store.get<string[]>(KEYS.unlinked, []);
  if (!unlinked.includes(emailLower)) {
    unlinked.push(emailLower);
    store.set(KEYS.unlinked, unlinked);
  }

  if ((window as any).supabaseProviderToken && removed.access_token === (window as any).supabaseProviderToken) {
    delete (window as any).supabaseProviderToken;
  }

  removed.access_token = undefined;
  removed.refresh_token = undefined;
  removed.token = undefined;
  removed.status = undefined;

  if (currentUser) {
    try {
      if (removed.id) {
        await supabase.from('user_accounts').delete().eq('id', removed.id);
      }
      await supabase
        .from('user_accounts')
        .delete()
        .eq('user_id', currentUser.id)
        .ilike('email', emailLower);

      if (accounts.length > 0 && accounts[0]) {
        await supabase
          .from('user_accounts')
          .update({ is_primary: true })
          .eq('user_id', currentUser.id)
          .ilike('email', accounts[0].email);
      }

      if (!options.fromBrowserClose && currentUser.identities && Array.isArray(currentUser.identities)) {
        const matchingIdentity = currentUser.identities.find(
          (id: any) => id.identity_data?.email?.toLowerCase() === emailLower || id.email?.toLowerCase() === emailLower
        );
        if (matchingIdentity) {
          await supabase.auth.unlinkIdentity(matchingIdentity);
        }
      }
    } catch (err) {
      console.error('Excepción al eliminar cuenta de Supabase:', err);
    }
  }

  if (!accounts.length) {
    accounts = [];
    mails = [];
    store.del(KEYS.accounts);
    store.del(KEYS.mails);
    sessionStorage.removeItem(KEYS.sessionAccounts);
  }

  persist();
  if (!options.fromBrowserClose) {
    hydrateProfileForm();
    renderAccounts();
    renderMails();
    populateComposeFromSelect();
    if (!options.silent) {
      toast(`✔ ${t('account_unlinked').replace('{email}', emailLower)}`);
    }
  }

  return true;
}

function unlinkAllAccountsOnBrowserClose(): void {
  if (accounts.length === 0 || !currentUser || !cachedAuthToken) return;

  const emailsToUnlink = accounts.map((a) => a.email.toLowerCase());
  deleteAccountsFromDbKeepalive(currentUser.id, cachedAuthToken, emailsToUnlink);

  store.set(KEYS.unlinkNotice, { emails: emailsToUnlink, at: Date.now() });

  for (const email of [...emailsToUnlink]) {
    const unlinked = store.get<string[]>(KEYS.unlinked, []);
    if (!unlinked.includes(email)) {
      unlinked.push(email);
      store.set(KEYS.unlinked, unlinked);
    }
  }

  delete (window as any).supabaseProviderToken;

  // Invalidar tokens en memoria; sessionStorage se limpia solo al cerrar el navegador
  accounts.forEach((a) => {
    a.access_token = undefined;
    a.refresh_token = undefined;
    a.token = undefined;
  });
  store.del(KEYS.accounts);
}

function inferProvider(email: string, token?: string): EmailProvider {
  if (token?.startsWith('ya29')) return 'gmail';
  const lower = email.toLowerCase();
  if (lower.includes('outlook') || lower.includes('hotmail') || lower.includes('live') || lower.includes('onmicrosoft.com')) {
    return 'outlook';
  }
  if (lower.includes('gmail') || lower.includes('googlemail')) return 'gmail';
  return 'custom';
}

async function resolveSendToken(acc: Account): Promise<string | null> {
  const { data: { session } } = await supabase.auth.getSession();
  if (session?.access_token) cachedAuthToken = session.access_token;

  const candidates: string[] = [];
  if (session?.provider_token && !session.provider_token.startsWith('eyJ')) {
    candidates.push(session.provider_token);
  }
  if (acc.access_token) candidates.push(acc.access_token);
  if (acc.token) candidates.push(acc.token);
  const globalToken = (window as any).supabaseProviderToken;
  if (globalToken) candidates.push(globalToken);

  for (const token of candidates) {
    if (!token || token.startsWith('eyJ')) continue;
    if (acc.provider === 'gmail' && token.startsWith('ya29')) return token;
    if (acc.provider === 'outlook') return token;
  }
  return null;
}

async function fetchGoogleUserEmail(accessToken: string): Promise<string | null> {
  try {
    const res = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (res.ok) {
      const data = await res.json();
      return data.email || null;
    }
  } catch (err) {
    console.warn('Could not fetch Google userinfo:', err);
  }
  return null;
}

async function fetchOutlookUserEmail(accessToken: string): Promise<string | null> {
  try {
    const res = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (res.ok) {
      const data = await res.json();
      return data.mail || data.userPrincipalName || null;
    }
  } catch (err) {
    console.warn('Could not fetch Outlook userinfo:', err);
  }
  return null;
}

async function handleDirectOAuthRedirect(): Promise<boolean> {
  const hash = window.location.hash || '';
  if (!hash.includes('access_token=') && !hash.includes('provider_token=')) return false;

  const params = new URLSearchParams(hash.replace(/^#/, ''));
  const providerToken = params.get('provider_token');
  const accessToken = params.get('access_token');

  const tokenToUse = (providerToken && providerToken.startsWith('ya29'))
    ? providerToken
    : (accessToken && !accessToken.startsWith('eyJ'))
      ? accessToken
      : providerToken;

  if (!tokenToUse) return false;

  (window as any).supabaseProviderToken = tokenToUse;
  let userEmail = await fetchGoogleUserEmail(tokenToUse);
  if (!userEmail) userEmail = await fetchOutlookUserEmail(tokenToUse);

  if (userEmail) {
    history.replaceState(null, '', window.location.pathname);
    // Remover de la lista de desvinculados explícitos si el usuario vuelve a vincular esta cuenta
    const unlinked = store.get<string[]>(KEYS.unlinked, []).filter((e) => e.toLowerCase() !== userEmail!.toLowerCase());
    store.set(KEYS.unlinked, unlinked);

    const provider: EmailProvider = userEmail.includes('outlook') || userEmail.includes('hotmail') || userEmail.includes('live') ? 'outlook' : 'gmail';
    const accIndex = accounts.findIndex((a) => a.email.toLowerCase() === userEmail!.toLowerCase());
    if (accIndex !== -1) {
      accounts[accIndex].token = tokenToUse;
      accounts[accIndex].access_token = tokenToUse;
      accounts[accIndex].status = 'connected';
    } else {
      accounts.push({
        email: userEmail.toLowerCase(),
        primary: accounts.length === 0,
        provider,
        status: 'connected',
        token: tokenToUse,
        access_token: tokenToUse,
      });
    }

    if (currentUser) {
      await supabase.from('user_accounts').upsert({
        user_id: currentUser.id,
        email: userEmail.toLowerCase(),
        access_token: tokenToUse,
        is_primary: accounts.length === 1,
      }, { onConflict: 'user_id,email' });
    }

    activeAccount = userEmail.toLowerCase();
    const count = await syncAccountInbox(userEmail, provider, tokenToUse);
    persist();
    renderAccounts();
    renderMails();
    toast(`✔ Cuenta ${userEmail} conectada con ${provider === 'gmail' ? 'Google' : 'Microsoft'} OAuth (${count} correos).`);
    return true;
  }
  return false;
}

async function syncSupabaseData(): Promise<void> {
  try {
    const { data: { session } } = await supabase.auth.getSession();
    if (!session || !session.user) {
      window.location.href = './';
      return;
    }
    currentUser = session.user;
    cachedAuthToken = session.access_token;

    const wasRecentReload = (() => {
      const unloadTime = sessionStorage.getItem('triade.unloadTime');
      sessionStorage.removeItem('triade.unloadTime');
      return unloadTime ? Date.now() - parseInt(unloadTime, 10) < 8000 : false;
    })();

    // Procesar primero cualquier vinculación directa por hash (#access_token=ya29...) para no alterar la sesión principal
    await handleDirectOAuthRedirect();

    const hashParams = new URLSearchParams((window.location.hash || '').replace(/^#/, ''));
    const hashProviderToken = hashParams.get('provider_token');
    const providerToken = session.provider_token || hashProviderToken || (window as any).supabaseProviderToken;

    if (providerToken) {
      (window as any).supabaseProviderToken = providerToken;
    }

    // 1. Cargar o inicializar Perfil (manteniendo siempre el correo principal de la sesión inicial)
    const { data: pData } = await supabase.from('profiles').select('*').eq('id', currentUser.id).single();
    if (pData && pData.email) {
      profile = {
        name: pData.name || currentUser.user_metadata?.full_name || 'Usuario Triade',
        email: pData.email,
        phone: pData.phone || '',
        signature: pData.signature || '— Enviado desde Triade Mail',
      };
    } else {
      profile.email = profile.email || session.user.email || currentUser.email || '';
      profile.name = profile.name || currentUser.user_metadata?.full_name || 'Usuario Triade';
    }

    // 2. Cargar Categorías desde la tabla de Supabase DB
    const { data: cData } = await supabase.from('categories').select('*').eq('user_id', currentUser.id);
    if (cData) {
      categories = cData.map((c: any) => ({ id: c.id, name: c.name, color: c.color, keywords: c.keywords || [] }));
    } else {
      categories = [];
    }

    // 3. Cargar Cuentas del usuario guardadas en Supabase (solo las conectadas por OAuth)
    const { data: dbAccounts } = await supabase.from('user_accounts').select('*').eq('user_id', currentUser.id);

    if (dbAccounts !== null && dbAccounts.length > 0) {
      const dbMapped = dbAccounts.map((a: any) => ({
        id: a.id,
        user_id: a.user_id,
        email: a.email.toLowerCase(),
        primary: a.is_primary,
        provider: a.email.toLowerCase().includes('gmail') ? 'gmail' as EmailProvider : a.email.toLowerCase().includes('outlook') ? 'outlook' as EmailProvider : 'custom' as EmailProvider,
        status: 'connected' as const,
        access_token: a.access_token || undefined,
        refresh_token: a.refresh_token || undefined,
        expires_at: a.expires_at || undefined,
        token: a.access_token || undefined,
      }));

      // Preferir tokens de sessionStorage (sesión activa) sobre los de la DB
      const sessionMap = new Map(accounts.map((a) => [a.email.toLowerCase(), a]));
      accounts = dbMapped.map((dbAcc) => {
        const sessionAcc = sessionMap.get(dbAcc.email.toLowerCase());
        if (sessionAcc?.access_token || sessionAcc?.token) {
          return { ...dbAcc, access_token: sessionAcc.access_token || sessionAcc.token, token: sessionAcc.token || sessionAcc.access_token };
        }
        return dbAcc;
      });

      // Añadir cuentas solo en sessionStorage que aún no están en DB
      for (const sessAcc of sessionMap.values()) {
        if (!accounts.some((a) => a.email.toLowerCase() === sessAcc.email.toLowerCase())) {
          accounts.push(sessAcc);
        }
      }
    } else if (accounts.length === 0) {
      accounts = [];
    }

    // Si retornó un providerToken de OAuth, obtener el correo real de la API de Google/Microsoft (no usar el login email)
    if (providerToken) {
      let oauthEmail = await fetchGoogleUserEmail(providerToken);
      if (!oauthEmail) {
        oauthEmail = await fetchOutlookUserEmail(providerToken);
      }

      if (oauthEmail) {
        const cleanEmail = oauthEmail.toLowerCase().trim();
        const unlinkedSet = new Set(store.get<string[]>(KEYS.unlinked, []));

        // Solo restaurar la cuenta si NO fue desvinculada explícitamente por el usuario
        if (!unlinkedSet.has(cleanEmail)) {
          const provider: EmailProvider = (cleanEmail.includes('outlook') || cleanEmail.includes('hotmail') || cleanEmail.includes('live')) ? 'outlook' : 'gmail';
          const existingAcc = accounts.find((a) => a.email.toLowerCase() === cleanEmail);

          if (!existingAcc) {
            accounts.push({
              email: cleanEmail,
              primary: accounts.length === 0,
              provider,
              status: 'connected',
              access_token: providerToken,
              token: providerToken,
            });
          } else {
            existingAcc.access_token = providerToken;
            existingAcc.token = providerToken;
            existingAcc.status = 'connected';
          }

          const { error: accErr } = await supabase.from('user_accounts').upsert({
            user_id: currentUser.id,
            email: cleanEmail,
            access_token: providerToken,
            refresh_token: session.provider_refresh_token || null,
            is_primary: accounts.length === 1,
          }, { onConflict: 'user_id,email' });

          if (accErr) console.warn('Error al guardar/actualizar user_accounts en Supabase:', accErr);
        }
      }
    }

    // 4. Cargar Correos guardados en Supabase
    const { data: dbMails, error: mailSelectErr } = await supabase.from('mails').select('*').eq('user_id', currentUser.id);
    if (mailSelectErr) {
      console.warn('Nota: La consulta a la tabla "mails" en Supabase devolvió un error (ej: RLS o tipo de columna id):', mailSelectErr.message);
    }
    if (dbMails && dbMails.length > 0) {
      const loadedMails: Mail[] = dbMails.map((m: any) => ({
        id: m.id,
        from: m.from_name || 'Remitente',
        fromEmail: m.from_email || '',
        to: m.to_address || '',
        subject: m.subject || '(Sin asunto)',
        body: m.body || '',
        bodyHtml: m.body_html || undefined,
        account: m.account,
        folder: m.folder || 'inbox',
        unread: Boolean(m.unread),
        starred: Boolean(m.starred),
        time: m.time_label || new Date(m.created_at).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
        timestamp: m.created_at ? new Date(m.created_at).getTime() : Date.now(),
        scheduledFor: m.scheduled_for || null,
      }));

      const mailMap = new Map<string, Mail>();
      loadedMails.forEach((m) => mailMap.set(m.id, m));
      mails.forEach((m) => {
        if (!mailMap.has(m.id)) mailMap.set(m.id, m);
      });
      mails = Array.from(mailMap.values());
    }

    // 5. Sincronizar bandejas para las cuentas de correo vinculadas en vivo
    if (accounts.length > 0) {
      const validAccountEmails = new Set(accounts.map((a) => a.email.toLowerCase()));
      mails = mails.filter((m) => m.account && validAccountEmails.has(m.account.toLowerCase()));

      if (wasRecentReload) {
        store.del(KEYS.unlinkNotice);
        await restoreAccountsToDb();
      }

      for (const acc of accounts) {
        const effToken = acc.access_token || acc.token || (window as any).supabaseProviderToken;
        if (effToken) {
          await syncAccountInbox(acc.email, acc.provider, effToken);
        }
      }
    } else if (!wasRecentReload) {
      showPendingUnlinkNotices();
    }

    persist();
    hydrateProfileForm();
    renderAccounts();
    renderMails();
    renderCategoriesEditor();
  } catch (err) {
    console.error('Error al sincronizar datos de Supabase:', err);
  }
}

// Escuchar cambios de autenticación (ej: redirección tras inicio de sesión con Google OAuth)
supabase.auth.onAuthStateChange(async (event, session) => {
  if (session && session.user) {
    cachedAuthToken = session.access_token;
    if (session.provider_token) {
      (window as any).supabaseProviderToken = session.provider_token;
    }
    await syncSupabaseData();
  }
});


// ---------- Classification / Filtering ----------
function classify(subject: string): string {
  const s = (subject || '').toLowerCase();
  for (const cat of categories) {
    if (cat.keywords.some((k) => k && s.includes(k.toLowerCase()))) return cat.id;
  }
  return '__uncat';
}

function folderMails(folder: string): Mail[] {
  return mails.filter((m) => {
    if (folder === 'starred') return m.starred && m.folder !== 'trash' && m.folder !== 'spam';
    return m.folder === folder;
  });
}

function baseFiltered(): Mail[] {
  return folderMails(activeFolder).filter((m) => {
    if (activeAccount !== 'all' && m.account !== activeAccount) return false;
    if (searchQuery) {
      const q = searchQuery.toLowerCase();
      if (!`${m.from} ${m.to || ''} ${m.subject} ${m.body}`.toLowerCase().includes(q)) return false;
    }
    return true;
  });
}

function tabFiltered(): Mail[] {
  const base = baseFiltered();
  const list = activeTab === 'all'
    ? base
    : activeTab === '__uncat'
      ? base.filter((m) => classify(m.subject) === '__uncat')
      : base.filter((m) => classify(m.subject) === activeTab);

  // Ordenar SIEMPRE de más reciente a más antiguo
  return list.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
}

function countMailsForTab(tabId: string): number {
  const base = baseFiltered();
  if (tabId === 'all') return base.length;
  if (tabId === '__uncat') return base.filter((m) => classify(m.subject) === '__uncat').length;
  return base.filter((m) => classify(m.subject) === tabId).length;
}

// ---------- Rendering ----------
function updateFolderCounts(): void {
  document.querySelectorAll<HTMLElement>('[data-count]').forEach((el) => {
    const f = el.dataset.count!;
    const n = folderMails(f).length;
    el.textContent = String(n);
    el.style.display = n ? '' : 'none';
  });
}

function escapeAttr(s: unknown): string {
  return String(s ?? '').replace(/"/g, '&quot;').replace(/</g, '&lt;');
}
function escapeHtml(s: unknown): string {
  return String(s ?? '').replace(
    /[&<>"']/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] as string)
  );
}

async function syncAccountInbox(accountEmailRaw: string, forcedProvider?: EmailProvider, token?: string): Promise<number> {
  const accountEmail = accountEmailRaw.trim().toLowerCase();
  if (!accountEmail || !accountEmail.includes('@')) return 0;

  let provider: EmailProvider = forcedProvider || 'custom';
  if (!forcedProvider || forcedProvider === 'custom') {
    if (accountEmail.includes('gmail.com')) provider = 'gmail';
    else if (accountEmail.includes('outlook.com') || accountEmail.includes('hotmail.com') || accountEmail.includes('live.com')) provider = 'outlook';
  }

  const accIndex = accounts.findIndex((a) => a.email.toLowerCase() === accountEmail);
  const nowTimeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

  if (accIndex === -1) {
    accounts.push({
      email: accountEmail,
      primary: accounts.length === 0,
      provider,
      status: 'connected',
      token,
      access_token: token,
      lastSync: nowTimeStr,
    });
  } else {
    accounts[accIndex].provider = provider;
    accounts[accIndex].status = 'connected';
    accounts[accIndex].lastSync = nowTimeStr;
    if (token) {
      accounts[accIndex].token = token;
      accounts[accIndex].access_token = token;
    }
  }

  const acc = accounts.find((a) => a.email.toLowerCase() === accountEmail);
  const effectiveToken = token || acc?.access_token || acc?.token || (window as any).supabaseProviderToken;
  let liveMails: Mail[] = [];

  if (effectiveToken) {
    try {
      if (provider === 'gmail') {
        liveMails = await fetchRealGmailMails(effectiveToken, accountEmail);
      } else if (provider === 'outlook') {
        liveMails = await fetchRealOutlookMails(effectiveToken, accountEmail);
      }
    } catch (err: any) {
      console.warn('Could not fetch live mails with token:', err);
      toast(`⚠️ Error al conectar con ${provider === 'gmail' ? 'Gmail' : 'Outlook'} API: Token inválido o no configurado en Supabase.`);
    }
  } else {
    toast(`⚠️ Se requiere un Token OAuth o API Key para extraer los correos en vivo de ${accountEmail}.`);
  }

  let addedCount = 0;
  if (liveMails.length > 0) {
    // Reemplazar correos antiguos de esta cuenta con los correos reales en vivo extraídos de la API (solo en memoria)
    mails = mails.filter((m) => m.account.toLowerCase() !== accountEmail.toLowerCase());
    mails.push(...liveMails);
    addedCount = liveMails.length;
  }

  // Ordenar SIEMPRE la lista completa de correos de más reciente a más antiguo
  mails.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));

  if (currentUser) {
    supabase.from('user_accounts').upsert({
      user_id: currentUser.id,
      email: accountEmail,
      access_token: effectiveToken || acc?.access_token || acc?.token || null,
      refresh_token: acc?.refresh_token || null,
      expires_at: acc?.expires_at || null,
      is_primary: acc?.primary || false,
    }, { onConflict: 'user_id,email' }).then(({ error }) => {
      if (error) console.warn('Error al actualizar cuenta en Supabase:', error);
    });
  }

  persist();
  return addedCount;
}

function renderAccounts(): void {
  const list = document.getElementById('accountsList')!;
  list.innerHTML = '';
  const chipAll = document.createElement('div');
  chipAll.className = 'acct-chip' + (activeAccount === 'all' ? ' active' : '');
  chipAll.innerHTML = `<span class="dot"></span><span>${escapeHtml(t('all'))}</span>`;
  chipAll.onclick = () => {
    activeAccount = 'all';
    currentPage = 1;
    renderAccounts();
    renderMails();
  };
  list.appendChild(chipAll);
  accounts.forEach((a) => {
    const el = document.createElement('div');
    el.className = 'acct-chip' + (activeAccount === a.email ? ' active' : '');
    const badge = a.provider === 'gmail' ? ' 🔴' : a.provider === 'outlook' ? ' 🔵' : '';
    el.innerHTML = `<span class="dot"></span><span>${escapeHtml(a.email)}${badge}</span>`;
    el.onclick = () => {
      activeAccount = a.email;
      currentPage = 1;
      renderAccounts();
      renderMails();
    };
    list.appendChild(el);
  });

  const edit = document.getElementById('accountsEdit');
  if (edit) {
    edit.innerHTML = '';
    accounts.forEach((a, i) => {
      const row = document.createElement('div');
      row.className = 'account-item';
      const providerLabel = a.provider === 'gmail' ? 'Gmail' : a.provider === 'outlook' ? 'Outlook' : 'Email';
      const syncInfo = a.lastSync ? `<span class="hint" style="font-size:0.75rem;margin-left:0.5rem">${t('last_synced')} ${a.lastSync}</span>` : '';
      row.innerHTML = `
        <div style="flex:1;min-width:0;display:flex;align-items:center;gap:0.5rem;flex-wrap:wrap">
          <b class="email">${escapeHtml(a.email)}</b>
          <span class="chip" style="font-size:0.7rem;padding:0.1rem 0.4rem;border-radius:4px;background:var(--accent-glow)">${providerLabel}</span>
          ${a.primary ? `<span class="badge">${escapeHtml(t('primary_account'))}</span>` : ''}
          ${syncInfo}
        </div>
        <button class="btn sm primary ghost" data-sync="${escapeAttr(a.email)}" title="${t('sync_inbox')}">🔄</button>
        <button class="btn sm ghost" data-remove="${i}">${escapeHtml(t('remove'))}</button>`;
      edit.appendChild(row);
    });
    edit.querySelectorAll<HTMLElement>('[data-sync]').forEach((b) => {
      b.onclick = async () => {
        const emailToSync = b.dataset.sync!;
        const acc = accounts.find((x) => x.email === emailToSync);
        const count = await syncAccountInbox(emailToSync, acc?.provider, acc?.access_token || acc?.token);
        renderAccounts();
        renderMails();
        toast(`${t('synced_success')} (${count} nuevos)`);
      };
    });
    edit.querySelectorAll<HTMLElement>('[data-remove]').forEach((b) => {
      b.onclick = async () => {
        const index = +b.dataset.remove!;
        const removed = accounts[index];
        if (!removed) return;
        await unlinkAccountByEmail(removed.email);
      };
    });
  }
}

function renderTabs(): void {
  const tabs = document.getElementById('tabs')!;
  tabs.innerHTML = '';
  const items: { id: string; name: string; color: string | null }[] = [
    { id: 'all', name: t('all'), color: null },
    ...categories.map((c) => ({ id: c.id, name: c.name, color: c.color })),
    { id: '__uncat', name: t('uncategorized'), color: '#6b7299' },
  ];
  items.forEach((it) => {
    const count = countMailsForTab(it.id);
    const el = document.createElement('div');
    el.className = 'tab' + (activeTab === it.id ? ' active' : '');
    const dot = it.color
      ? `<span style="width:8px;height:8px;border-radius:50%;background:${it.color};box-shadow:0 0 6px ${it.color};display:inline-block"></span>`
      : '';
    el.innerHTML = `${dot}<span>${escapeHtml(it.name)}</span>${count ? `<span class="badge">${count}</span>` : ''}`;
    el.onclick = () => {
      activeTab = it.id;
      currentPage = 1;
      renderTabs();
      renderMails();
    };
    tabs.appendChild(el);
  });
}

function emptyMessageForFolder(): string {
  return t('empty_' + activeFolder) || t('no_mails');
}

function renderMails(): void {
  renderTabs();
  updateFolderCounts();
  const list = document.getElementById('mailList')!;

  if (!accounts.length) {
    list.innerHTML = `
      <div class="empty-state" style="padding: 3rem 1.5rem; text-align: center;">
        <svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="var(--accent)" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" style="margin-bottom: 1rem; opacity: 0.85;"><rect width="20" height="16" x="2" y="4" rx="2"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/></svg>
        <h3 style="margin-bottom:0.5rem">No has registrado ninguna cuenta de correo</h3>
        <p class="hint" style="max-width: 420px; margin: 0 auto 1.5rem auto;">Conecta tu correo de Gmail u Outlook para obtener todos tus mensajes en la bandeja centralizada.</p>
        <button class="btn primary" id="emptyStateConnectBtn">Registrar correo de Gmail u Outlook</button>
      </div>`;
    const btn = document.getElementById('emptyStateConnectBtn');
    if (btn) {
      btn.onclick = () => {
        document.getElementById('connectProviderModal')?.classList.add('open');
      };
    }
    renderPagination(0, 1);
    return;
  }

  const all = tabFiltered();
  const total = all.length;
  const totalPages = Math.max(1, Math.ceil(total / pageSize));
  if (currentPage > totalPages) currentPage = totalPages;
  const start = (currentPage - 1) * pageSize;
  const items = all.slice(start, start + pageSize);

  list.innerHTML = '';
  if (!items.length) {
    list.innerHTML = `<div class="empty-state">${escapeHtml(emptyMessageForFolder())}</div>`;
  } else {
    items.forEach((m) => {
      const catId = classify(m.subject);
      const cat = categories.find((c) => c.id === catId);
      const chipColor = cat ? cat.color : '#6b7299';
      const chipName = cat ? cat.name : t('uncategorized');
      const isSentLike = m.folder === 'sent' || m.folder === 'drafts';
      const primaryLine = isSentLike && m.to ? m.to : m.from || m.to || '—';

      const row = document.createElement('div');
      row.className = 'mail-row' + (m.unread ? ' unread' : '') + (m.starred ? ' starred' : '');
      row.innerHTML = `
        <button class="star" title="${escapeAttr(t(m.starred ? 'unstar' : 'star'))}" data-action="star">${m.starred ? '★' : '☆'}</button>
        <div class="from">${escapeHtml(primaryLine)}</div>
        <div class="body-wrap">
          <span class="chip" style="background:${chipColor}22;border-color:${chipColor}66;color:var(--fg)">${escapeHtml(chipName)}</span>
          <b>${escapeHtml(m.subject || '(sin asunto)')}</b> <span class="body-sep">—</span> <span class="body">${escapeHtml(m.body || '')}</span>
        </div>
        <div class="time">${escapeHtml(m.time || '')}</div>
        <div class="row-actions">
          ${
            m.folder === 'trash'
              ? `<button class="row-btn" data-action="restore" title="${escapeAttr(t('restore'))}">↺</button>
                 <button class="row-btn" data-action="delete" title="${escapeAttr(t('delete_forever'))}">✕</button>`
              : m.folder === 'spam'
                ? `<button class="row-btn" data-action="restore" title="${escapeAttr(t('restore'))}">↺</button>
                   <button class="row-btn" data-action="trash" title="${escapeAttr(t('move_trash'))}">🗑</button>`
                : `<button class="row-btn" data-action="spam" title="${escapeAttr(t('move_spam'))}">⚠</button>
                   <button class="row-btn" data-action="trash" title="${escapeAttr(t('move_trash'))}">🗑</button>`
          }
        </div>`;
      if (m.scheduledFor) {
        const bw = row.querySelector('.body-wrap')!;
        const badge = document.createElement('span');
        badge.className = 'status-chip';
        badge.textContent = t('scheduled');
        bw.appendChild(badge);
      }
      row.querySelectorAll<HTMLElement>('[data-action]').forEach((btn) => {
        btn.addEventListener('click', (e) => {
          e.stopPropagation();
          handleMailAction(m.id, btn.dataset.action!);
        });
      });
      row.addEventListener('click', () => {
        if (m.folder === 'drafts') {
          openCompose(m);
          return;
        }
        if (m.unread) {
          m.unread = false;
          persist();
        }
        openReader(m);
      });
      list.appendChild(row);
    });
  }
  renderPagination(total, totalPages);
}

function handleMailAction(id: string, action: string): void {
  void handleMailActionAsync(id, action);
}

async function handleMailActionAsync(id: string, action: string): Promise<void> {
  const m = mails.find((x) => x.id === id);
  if (!m) return;
  if (action === 'star') {
    m.starred = !m.starred;
  } else if (action === 'spam') {
    m.folder = 'spam';
    m.starred = false;
  } else if (action === 'trash') {
    m.folder = 'trash';
    m.starred = false;
  } else if (action === 'restore') {
    m.folder = 'inbox';
  } else if (action === 'delete') {
    const confirmed = await showConfirm({
      title: t('delete_forever'),
      message: t('confirm_delete'),
      confirmText: t('delete_forever'),
      danger: true,
    });
    if (!confirmed) return;
    const idx = mails.findIndex((x) => x.id === id);
    if (idx >= 0) mails.splice(idx, 1);
    toast(t('mail_deleted'));
  }
  persist();
  renderMails();
}

function renderPagination(total: number, totalPages: number): void {
  let el = document.getElementById('pagination');
  if (!el) {
    el = document.createElement('div');
    el.className = 'pagination';
    el.id = 'pagination';
    document.querySelector('.main')!.appendChild(el);
  }
  const from = total === 0 ? 0 : (currentPage - 1) * pageSize + 1;
  const to = Math.min(currentPage * pageSize, total);
  const pages: number[] = [];
  const win = 1;
  pages.push(1);
  for (let p = Math.max(2, currentPage - win); p <= Math.min(totalPages - 1, currentPage + win); p++) {
    if (p !== 1 && p !== totalPages) pages.push(p);
  }
  if (totalPages > 1) pages.push(totalPages);
  const rendered: (number | string)[] = [];
  let prev = 0;
  pages.forEach((p) => {
    if (p - prev > 1) rendered.push('…');
    rendered.push(p);
    prev = p;
  });

  el.innerHTML = `
    <div class="info">${escapeHtml(t('showing'))} <b>${from}–${to}</b> ${escapeHtml(t('of'))} <b>${total}</b></div>
    <div class="controls">
      <select class="page-size" id="pageSizeSel" aria-label="per page">
        ${[10, 25, 50, 100].map((n) => `<option value="${n}" ${n === pageSize ? 'selected' : ''}>${n} ${escapeHtml(t('per_page'))}</option>`).join('')}
      </select>
      <button class="page-btn" id="prevPage" ${currentPage <= 1 ? 'disabled' : ''}>‹</button>
      ${rendered
        .map((p) =>
          p === '…'
            ? `<span class="page-btn ellipsis">…</span>`
            : `<button class="page-btn ${p === currentPage ? 'active' : ''}" data-page="${p}">${p}</button>`
        )
        .join('')}
      <button class="page-btn" id="nextPage" ${currentPage >= totalPages ? 'disabled' : ''}>›</button>
    </div>`;

  el.querySelector<HTMLElement>('#prevPage')!.onclick = () => {
    if (currentPage > 1) {
      currentPage--;
      renderMails();
    }
  };
  el.querySelector<HTMLElement>('#nextPage')!.onclick = () => {
    if (currentPage < totalPages) {
      currentPage++;
      renderMails();
    }
  };
  el.querySelectorAll<HTMLElement>('[data-page]').forEach((b) => {
    b.onclick = () => {
      currentPage = +b.dataset.page!;
      renderMails();
    };
  });
  el.querySelector<HTMLSelectElement>('#pageSizeSel')!.onchange = (e) => {
    pageSize = +(e.target as HTMLSelectElement).value;
    currentPage = 1;
    persist();
    renderMails();
  };
}

// ---------- Categories editor ----------
function renderCategoriesEditor(): void {
  const wrap = document.getElementById('categoriesEdit');
  if (!wrap) return;
  wrap.innerHTML = '';
  if (!categories.length) {
    wrap.innerHTML = `<p class="hint">${escapeHtml(t('no_categories'))}</p>`;
    return;
  }
  categories.forEach((cat, idx) => {
    const card = document.createElement('div');
    card.className = 'category-card';
    card.innerHTML = `
      <div class="category-head">
        <span class="swatch-mini" style="background:${cat.color}; color:${cat.color}"></span>
        <input type="text" value="${escapeAttr(cat.name)}" data-name />
        <input type="color" value="${cat.color}" data-color />
        <button class="del" data-del title="${escapeAttr(t('remove'))}">✕</button>
      </div>
      <div class="category-kws"></div>
      <div class="category-add">
        <input type="text" data-newkw placeholder="${escapeAttr(t('new_keyword_placeholder'))}" />
        <button data-addkw>${escapeHtml(t('add'))}</button>
      </div>`;
    const kws = card.querySelector('.category-kws')!;
    cat.keywords.forEach((kw, kIdx) => {
      const tag = document.createElement('span');
      tag.className = 'tag';
      tag.innerHTML = `${escapeHtml(kw)} <button title="remove">✕</button>`;
      tag.querySelector('button')!.onclick = () => {
        categories[idx].keywords.splice(kIdx, 1);
        persist();
        renderCategoriesEditor();
        renderMails();
      };
      kws.appendChild(tag);
    });
    (card.querySelector('[data-name]') as HTMLInputElement).oninput = (e) => {
      categories[idx].name = (e.target as HTMLInputElement).value;
      persist();
      renderTabs();
    };
    (card.querySelector('[data-color]') as HTMLInputElement).oninput = (e) => {
      categories[idx].color = (e.target as HTMLInputElement).value;
      const sm = card.querySelector<HTMLElement>('.swatch-mini')!;
      sm.style.background = (e.target as HTMLInputElement).value;
      sm.style.color = (e.target as HTMLInputElement).value;
      persist();
      renderTabs();
      renderMails();
    };
    card.querySelector<HTMLElement>('[data-del]')!.onclick = () => {
      const [removed] = categories.splice(idx, 1);
      if (removed && currentUser) {
        supabase.from('categories').delete().eq('user_id', currentUser.id).eq('name', removed.name).then(({ error }) => {
          if (error) console.warn('Error al eliminar categoría en Supabase:', error);
        });
      }
      if (activeTab === cat.id) activeTab = 'all';
      persist();
      renderCategoriesEditor();
      renderMails();
    };
    const kwInput = card.querySelector<HTMLInputElement>('[data-newkw]')!;
    const addKw = () => {
      const v = kwInput.value.trim();
      if (!v) return;
      categories[idx].keywords.push(v);
      kwInput.value = '';
      persist();
      renderCategoriesEditor();
      renderMails();
      const cards = document.querySelectorAll('.category-card');
      (cards[idx] as HTMLElement | undefined)?.querySelector<HTMLInputElement>('[data-newkw]')?.focus();
    };
    card.querySelector<HTMLElement>('[data-addkw]')!.onclick = addKw;
    kwInput.onkeydown = (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        addKw();
      }
    };
    wrap.appendChild(card);
  });
}

function hydrateProfileForm(): void {
  (document.getElementById('profName') as HTMLInputElement).value = profile.name || '';
  (document.getElementById('profEmail') as HTMLInputElement).value = profile.email || '';
  (document.getElementById('profPhone') as HTMLInputElement).value = profile.phone || '';
  (document.getElementById('profSig') as HTMLTextAreaElement).value = profile.signature || '';
  const initial = (profile.name || 'T').charAt(0).toUpperCase();
  document.getElementById('userAvatar')!.textContent = initial;
  const alg = document.getElementById('userAvatarLg');
  if (alg) alg.textContent = initial;
  const nm = document.getElementById('userMenuName');
  if (nm) nm.textContent = profile.name || '—';
  const em = document.getElementById('userMenuEmail');
  if (em) em.textContent = profile.email || '—';
}

function hydrateAppearance(): void {
  const theme = store.get<string>(KEYS.theme, 'dark');
  const accent = store.get<string>(KEYS.accent, '#e63946');
  const lang = store.get<string>(KEYS.lang, 'es');
  document.querySelectorAll<HTMLElement>('.theme-card').forEach((c) =>
    c.classList.toggle('active', c.dataset.theme === theme)
  );
  document.querySelectorAll<HTMLElement>('.swatch[data-color]').forEach((s) =>
    s.classList.toggle('active', s.dataset.color === accent)
  );
  document.querySelectorAll<HTMLElement>('.lang-btn').forEach((b) =>
    b.classList.toggle('active', b.dataset.lang === lang)
  );
  (document.getElementById('customColor') as HTMLInputElement).value = accent;
}

// ---------- Toast ----------
let toastTimer: ReturnType<typeof setTimeout> | undefined;
function toast(msg: string, duration = 2200): void {
  const el = document.getElementById('toast')!;
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), duration);
}

// ---------- Confirm modal ----------
type ConfirmOptions = {
  title?: string;
  message: string;
  confirmText?: string;
  cancelText?: string;
  danger?: boolean;
};

let confirmResolver: ((value: boolean) => void) | null = null;

function closeConfirmModal(): void {
  document.getElementById('confirmModal')?.classList.remove('open');
  confirmResolver = null;
}

function initConfirmModal(): void {
  const modal = document.getElementById('confirmModal')!;
  const okBtn = document.getElementById('confirmOk') as HTMLButtonElement;
  const cancelBtn = document.getElementById('confirmCancel') as HTMLButtonElement;

  const finishConfirm = (confirmed: boolean) => {
    const resolver = confirmResolver;
    closeConfirmModal();
    if (resolver) resolver(confirmed);
  };

  okBtn.onclick = () => finishConfirm(true);
  cancelBtn.onclick = () => finishConfirm(false);
  modal.addEventListener('click', (e) => {
    if ((e.target as HTMLElement).id === 'confirmModal') finishConfirm(false);
  });
}

function showConfirm(opts: ConfirmOptions): Promise<boolean> {
  return new Promise((resolve) => {
    confirmResolver = resolve;

    const modal = document.getElementById('confirmModal')!;
    const titleEl = document.getElementById('confirmTitle')!;
    const messageEl = document.getElementById('confirmMessage')!;
    const okBtn = document.getElementById('confirmOk') as HTMLButtonElement;
    const cancelBtn = document.getElementById('confirmCancel') as HTMLButtonElement;

    titleEl.textContent = opts.title || t('confirm');
    messageEl.textContent = opts.message;
    okBtn.textContent = opts.confirmText || t('confirm');
    cancelBtn.textContent = opts.cancelText || t('cancel');
    okBtn.classList.toggle('danger', !!opts.danger);
    okBtn.classList.toggle('primary', !opts.danger);

    modal.classList.add('open');
    setTimeout(() => okBtn.focus(), 60);
  });
}

// ---------- Compose ----------
let composingDraftId: string | null = null;
let composeAttachments: Attachment[] = [];
let scheduledFor: string | null = null;

type DraftLike = Partial<Mail> & { id?: string };

function populateComposeFromSelect(): void {
  const sel = document.getElementById('composeFrom') as HTMLSelectElement | null;
  if (!sel) return;
  sel.innerHTML = '';
  if (!accounts.length) {
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = t('no_linked_accounts');
    sel.appendChild(opt);
    sel.disabled = true;
    composeFromAccount = '';
    return;
  }
  sel.disabled = false;
  accounts.forEach((a) => {
    const opt = document.createElement('option');
    opt.value = a.email;
    const providerLabel = a.provider === 'gmail' ? 'Gmail' : a.provider === 'outlook' ? 'Outlook' : 'Email';
    opt.textContent = `${a.email} (${providerLabel})${a.primary ? ' ★' : ''}`;
    sel.appendChild(opt);
  });
  const preferred =
    composeFromAccount && accounts.some((a) => a.email === composeFromAccount)
      ? composeFromAccount
      : activeAccount !== 'all' && accounts.some((a) => a.email === activeAccount)
        ? activeAccount
        : accounts.find((a) => a.primary)?.email || accounts[0]?.email || '';
  sel.value = preferred;
  composeFromAccount = preferred;
  sel.onchange = () => {
    composeFromAccount = sel.value;
  };
}

function openCompose(draft?: DraftLike): void {
  const modal = document.getElementById('composeModal')!;
  composingDraftId = draft?.id ?? null;
  composeAttachments = draft?.attachments ? [...draft.attachments] : [];
  scheduledFor = draft?.scheduledFor ?? null;
  populateComposeFromSelect();
  if (draft?.account && accounts.some((a) => a.email === draft.account)) {
    composeFromAccount = draft.account;
    const sel = document.getElementById('composeFrom') as HTMLSelectElement;
    if (sel) sel.value = draft.account;
  }
  (document.getElementById('composeTo') as HTMLInputElement).value = draft?.to || '';
  (document.getElementById('composeSubject') as HTMLInputElement).value = draft?.subject || '';
  document.getElementById('composeBody')!.innerHTML = draft?.bodyHtml || (draft?.body ? escapeHtml(draft.body) : '');
  updateComposeStatus();
  renderAttachments();
  modal.classList.add('open');
  setTimeout(() => document.getElementById('composeTo')!.focus(), 80);
}

function closeCompose(): void {
  document.getElementById('composeModal')!.classList.remove('open');
  document.getElementById('emojiPop')!.classList.remove('open');
  document.getElementById('schedulePop')!.classList.remove('open');
  composingDraftId = null;
  composeAttachments = [];
  scheduledFor = null;
}

function updateComposeStatus(): void {
  const el = document.getElementById('composeStatus')!;
  if (scheduledFor) {
    const d = new Date(scheduledFor);
    el.textContent = `${t('scheduled_for')} ${d.toLocaleString()}`;
  } else el.textContent = '';
}

function renderAttachments(): void {
  const wrap = document.getElementById('attachList')!;
  wrap.innerHTML = '';
  composeAttachments.forEach((a, i) => {
    const el = document.createElement('span');
    el.className = 'attach-item';
    el.innerHTML = `📎 ${escapeHtml(a.name)} <span class="kb">${a.size ? Math.round(a.size / 1024) + ' KB' : ''}</span> <button title="remove">✕</button>`;
    el.querySelector('button')!.onclick = () => {
      composeAttachments.splice(i, 1);
      renderAttachments();
    };
    wrap.appendChild(el);
  });
}

function readComposeFields(): { to: string; subject: string; bodyHtml: string; body: string } {
  const body = document.getElementById('composeBody')!;
  return {
    to: (document.getElementById('composeTo') as HTMLInputElement).value.trim(),
    subject: (document.getElementById('composeSubject') as HTMLInputElement).value.trim(),
    bodyHtml: body.innerHTML,
    body: body.innerText,
  };
}

async function sendMail(): Promise<void> {
  const f = readComposeFields();
  if (!f.to) {
    document.getElementById('composeTo')!.focus();
    return;
  }

  const fromSel = document.getElementById('composeFrom') as HTMLSelectElement;
  const acct = (fromSel?.value || composeFromAccount || '').trim().toLowerCase();
  if (!acct || !accounts.some((a) => a.email.toLowerCase() === acct)) {
    toast(`⚠️ ${t('send_error_no_account')}`);
    fromSel?.focus();
    return;
  }

  const targetAcc = accounts.find((a) => a.email.toLowerCase() === acct)!;
  const effToken = await resolveSendToken(targetAcc);
  const sendProvider = inferProvider(targetAcc.email, effToken || undefined) !== 'custom'
    ? inferProvider(targetAcc.email, effToken || undefined)
    : targetAcc.provider || 'custom';

  if (!effToken) {
    toast(`⚠️ ${t('send_error_no_token')}`);
    return;
  }

  // Actualizar token en la cuenta si obtuvimos uno más reciente
  targetAcc.access_token = effToken;
  targetAcc.token = effToken;
  persist();

  if (composingDraftId) {
    const idx = mails.findIndex((x) => x.id === composingDraftId);
    if (idx >= 0) mails.splice(idx, 1);
  }

  const now = new Date();
  const timeStr = scheduledFor
    ? new Date(scheduledFor).toLocaleString([], { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })
    : `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

  let sentOk = false;
  const bodyContent = f.bodyHtml || f.body;
  const providerLabel = sendProvider === 'gmail' ? 'Gmail' : sendProvider === 'outlook' ? 'Outlook' : 'Email';

  if (!scheduledFor) {
    toast(`📤 Enviando correo vía ${providerLabel}…`);
    let result: SendResult;
    if (sendProvider === 'gmail') {
      result = await sendRealGmailMail(effToken, f.to, f.subject || '(Sin asunto)', bodyContent);
    } else if (sendProvider === 'outlook') {
      result = await sendRealOutlookMail(effToken, f.to, f.subject || '(Sin asunto)', bodyContent);
    } else {
      toast(`⚠️ ${t('send_error_no_account')}`);
      return;
    }

    if (!result.ok) {
      const msg = result.message || t('send_error_failed');
      toast(`⚠️ ${msg}`, 5000);
      if (result.code === 'missing_scope' || result.code === 'expired_token' || result.code === 'invalid_token') {
        const relink = await showConfirm({
          title: t('manage_accounts'),
          message: `${msg}\n\n${t('confirm_relink')}`,
          confirmText: t('connect_sync'),
        });
        if (relink) {
          closeCompose();
          document.getElementById('settingsModal')?.classList.add('open');
          document.querySelector<HTMLElement>('.modal-nav .nav-item[data-section="parameters"]')?.click();
        }
      }
      return;
    }
    sentOk = true;
    toast(`✔ Correo enviado exitosamente desde ${targetAcc.email}`);
  } else {
    toast(t('mail_scheduled'));
    sentOk = true;
  }

  const newMailObj: Mail = {
    id: uid('m'),
    from: profile.name || targetAcc.email,
    fromEmail: targetAcc.email,
    to: f.to,
    subject: f.subject || '(sin asunto)',
    body: (f.body || '').split('\n')[0].slice(0, 180),
    bodyHtml: f.bodyHtml,
    attachments: composeAttachments.slice(),
    account: targetAcc.email,
    time: timeStr,
    timestamp: Date.now(),
    unread: false,
    starred: false,
    folder: scheduledFor ? 'drafts' : 'sent',
    scheduledFor: scheduledFor || null,
  };

  mails.unshift(newMailObj);

  if (currentUser && sentOk) {
    supabase.from('mails').insert({
      user_id: currentUser.id,
      from_name: newMailObj.from,
      from_email: newMailObj.fromEmail,
      to_address: newMailObj.to,
      subject: newMailObj.subject,
      body: newMailObj.body,
      body_html: newMailObj.bodyHtml,
      account: newMailObj.account,
      folder: newMailObj.folder,
      unread: false,
      starred: false,
      time_label: newMailObj.time,
    }).then(({ error }) => {
      if (error) console.warn('Error al guardar correo enviado en Supabase:', error);
    });
  }

  persist();
  closeCompose();
  renderMails();
}

function saveDraft(): void {
  const f = readComposeFields();
  if (!f.to && !f.subject && !f.body.trim() && !composeAttachments.length) {
    closeCompose();
    return;
  }
  const now = new Date();
  const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
  const acct = (document.getElementById('composeFrom') as HTMLSelectElement)?.value
    || composeFromAccount
    || accounts.find((a) => a.primary)?.email
    || accounts[0]?.email
    || profile.email;
  const payload = {
    to: f.to,
    subject: f.subject || '(sin asunto)',
    body: (f.body || '').split('\n')[0].slice(0, 180),
    bodyHtml: f.bodyHtml,
    attachments: composeAttachments.slice(),
    time: timeStr,
    scheduledFor: scheduledFor || null,
  };
  if (composingDraftId) {
    const m = mails.find((x) => x.id === composingDraftId);
    if (m) Object.assign(m, payload);
  } else {
    mails.unshift({
      id: uid('m'),
      from: profile.name || acct,
      fromEmail: acct,
      account: acct,
      unread: false,
      starred: false,
      folder: 'drafts' as Folder,
      ...payload,
    });
  }
  persist();
  toast(t('draft_saved'));
  closeCompose();
  renderMails();
}

function discardCompose(): void {
  if (composingDraftId) {
    const idx = mails.findIndex((x) => x.id === composingDraftId);
    if (idx >= 0) {
      mails.splice(idx, 1);
      persist();
      renderMails();
    }
  }
  closeCompose();
}

// ---------- Reader ----------
function openReader(m: Mail): void {
  const modal = document.getElementById('readerModal')!;
  document.getElementById('readerSubject')!.textContent = m.subject || '(sin asunto)';
  document.getElementById('readerFromName')!.textContent = m.from || '—';
  document.getElementById('readerFromEmail')!.textContent = m.fromEmail || m.to || '';
  document.getElementById('readerTime')!.textContent = m.time || '';
  document.getElementById('readerAvatar')!.textContent = (m.from || '?').charAt(0).toUpperCase();
  const body = document.getElementById('readerBody')!;
  let html = m.bodyHtml || escapeHtml(m.body || '');
  if (m.attachments && m.attachments.length) {
    html +=
      '<div style="margin-top:1rem;padding-top:.8rem;border-top:1px solid var(--border)"><b>' +
      escapeHtml(t('attachment')) +
      ':</b><div class="attach-list" style="padding:.4rem 0">' +
      m.attachments.map((a) => `<span class="attach-item">📎 ${escapeHtml(a.name)}</span>`).join('') +
      '</div></div>';
  }
  body.innerHTML = html;
  const starBtn = document.getElementById('readerStar')!;
  starBtn.textContent = m.starred ? '★' : '☆';
  starBtn.onclick = () => {
    m.starred = !m.starred;
    persist();
    starBtn.textContent = m.starred ? '★' : '☆';
    renderMails();
  };
  document.getElementById('readerTrash')!.onclick = () => {
    m.folder = 'trash';
    m.starred = false;
    persist();
    renderMails();
    modal.classList.remove('open');
  };
  document.getElementById('readerReply')!.onclick = () => {
    modal.classList.remove('open');
    openCompose({
      to: m.fromEmail || m.from,
      subject: 'Re: ' + (m.subject || ''),
      bodyHtml:
        '<br><br><blockquote style="border-left:3px solid var(--border);padding-left:.6rem;color:var(--fg-mute)">' +
        (m.bodyHtml || escapeHtml(m.body || '')) +
        '</blockquote>',
    });
  };
  document.getElementById('readerForward')!.onclick = () => {
    modal.classList.remove('open');
    openCompose({
      to: '',
      subject: 'Fwd: ' + (m.subject || ''),
      bodyHtml: '<br><br>--- ' + t('forward') + ' ---<br>' + (m.bodyHtml || escapeHtml(m.body || '')),
      attachments: m.attachments || [],
    });
  };
  modal.classList.add('open');
}

// ---------- Sidebar / mobile ----------
function toggleSidebar(force?: boolean): void {
  const sb = document.getElementById('sidebar')!;
  const sc = document.getElementById('sidebarScrim')!;
  const on = typeof force === 'boolean' ? force : !sb.classList.contains('open');
  sb.classList.toggle('open', on);
  sc.classList.toggle('open', on);
}

// ---------- Emoji ----------
const EMOJIS = ['😀','😃','😄','😁','😆','😅','🤣','😂','🙂','🙃','😉','😊','😇','🥰','😍','🤩','😘','😗','☺️','😚','😙','🥲','😋','😛','😜','🤪','😝','🤑','🤗','🤭','🤫','🤔','🤐','🤨','😐','😑','😶','😏','😒','🙄','😬','🤥','😌','😔','😪','🤤','😴','😷','🤒','🤕','🤢','🤮','🥵','🥶','🥴','😵','🤯','🤠','🥳','😎','🤓','🧐','😕','😟','🙁','☹️','😮','😯','😲','😳','🥺','😦','😧','😨','😰','😥','😢','😭','😱','😖','😣','😞','😓','😩','😫','🥱','😤','😡','😠','🤬','😈','👍','👎','👏','🙏','🤝','💪','🎉','🔥','✨','⭐','❤️','💔','💯','✅','❌'];

function buildEmojiGrid(): void {
  const g = document.getElementById('emojiGrid')!;
  g.innerHTML = EMOJIS.map((e) => `<button type="button" data-emoji="${e}">${e}</button>`).join('');
  g.querySelectorAll<HTMLButtonElement>('button').forEach((b) => {
    b.onclick = () => {
      insertAtEditor(b.dataset.emoji!);
      document.getElementById('emojiPop')!.classList.remove('open');
    };
  });
}

function insertAtEditor(html: string): void {
  const ed = document.getElementById('composeBody')!;
  ed.focus();
  document.execCommand('insertHTML', false, html);
}

function updateToolbarState(): void {
  const editor = document.getElementById('composeBody');
  if (!editor) return;
  document.querySelectorAll<HTMLElement>('#composeToolbar .tb-btn[data-cmd]').forEach((btn) => {
    const cmd = btn.dataset.cmd!;
    let active = false;
    try {
      active = document.queryCommandState(cmd);
    } catch {
      // ignore
    }
    btn.classList.toggle('active', !!active);
  });
  try {
    const fontName = document.queryCommandValue('fontName') || '';
    const ff = document.getElementById('fontFamilySel') as HTMLSelectElement | null;
    if (ff) {
      const val = fontName.replace(/"/g, '').trim();
      Array.from(ff.options).forEach(
        (o) =>
          (o.selected = !!val && !!o.value && val.toLowerCase().includes(o.value.split(',')[0].toLowerCase().replace(/'/g, '')))
      );
    }
    const size = document.queryCommandValue('fontSize');
    const fs = document.getElementById('fontSizeSel') as HTMLSelectElement | null;
    if (fs && size) Array.from(fs.options).forEach((o) => (o.selected = o.value === size));
  } catch {
    // ignore
  }
}

// ---------- Init ----------
document.addEventListener('DOMContentLoaded', async () => {
  applyI18n();
  initConfirmModal();
  renderAccounts();
  renderMails();
  renderCategoriesEditor();
  hydrateProfileForm();
  hydrateAppearance();

  const isOAuthRedirect = await handleDirectOAuthRedirect();
  if (!isOAuthRedirect) {
    await syncSupabaseData();
  }


  document.querySelectorAll<HTMLElement>('.sidebar .nav-item[data-folder]').forEach((n) => {
    n.onclick = () => {
      document.querySelectorAll('.sidebar .nav-item[data-folder]').forEach((x) => x.classList.remove('active'));
      n.classList.add('active');
      activeFolder = n.dataset.folder!;
      activeTab = 'all';
      currentPage = 1;
      renderMails();
      if (window.innerWidth < 860) toggleSidebar(false);
    };
  });

  document.getElementById('searchInput')!.addEventListener('input', (e) => {
    searchQuery = (e.target as HTMLInputElement).value;
    currentPage = 1;
    renderMails();
  });

  document.getElementById('menuToggle')!.onclick = () => toggleSidebar();
  document.getElementById('sidebarScrim')!.onclick = () => toggleSidebar(false);

  document.getElementById('composeBtn')!.onclick = () => openCompose();
  document.getElementById('closeCompose')!.onclick = closeCompose;
  document.getElementById('sendMail')!.onclick = sendMail;
  document.getElementById('saveDraft')!.onclick = saveDraft;
  document.getElementById('discardMail')!.onclick = discardCompose;
  document.getElementById('composeModal')!.addEventListener('click', (e) => {
    if ((e.target as HTMLElement).id === 'composeModal') closeCompose();
  });

  const editor = document.getElementById('composeBody')!;
  document.querySelectorAll<HTMLElement>('#composeToolbar .tb-btn[data-cmd]').forEach((b) => {
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.onclick = () => {
      editor.focus();
      document.execCommand(b.dataset.cmd!, false, undefined);
      updateToolbarState();
    };
  });
  (document.getElementById('fontFamilySel') as HTMLSelectElement).onchange = (e) => {
    editor.focus();
    document.execCommand('fontName', false, (e.target as HTMLSelectElement).value || 'sans-serif');
    updateToolbarState();
  };
  (document.getElementById('fontSizeSel') as HTMLSelectElement).onchange = (e) => {
    editor.focus();
    document.execCommand('fontSize', false, (e.target as HTMLSelectElement).value);
    updateToolbarState();
  };
  document.getElementById('fontColorPick')!.addEventListener('input', (e) => {
    editor.focus();
    document.execCommand('foreColor', false, (e.target as HTMLInputElement).value);
  });
  ['keyup', 'mouseup', 'selectionchange'].forEach((evt) => {
    if (evt === 'selectionchange') document.addEventListener(evt, updateToolbarState);
    else editor.addEventListener(evt, updateToolbarState);
  });

  const linkModal = document.getElementById('linkModal')!;
  const linkUrlInput = document.getElementById('linkUrlInput') as HTMLInputElement;
  const linkTextInput = document.getElementById('linkTextInput') as HTMLInputElement;
  const linkModalTitle = linkModal.querySelector('.link-header h3')!;
  const linkInsertConfirmBtn = document.getElementById('insertLinkConfirm') as HTMLButtonElement;
  let linkInsertMode: 'link' | 'drive' = 'link';
  const closeLinkModal = () => linkModal.classList.remove('open');

  document.getElementById('insertLinkBtn')!.onclick = () => {
    linkInsertMode = 'link';
    linkModalTitle.textContent = t('insert_link');
    linkInsertConfirmBtn.textContent = t('insert');
    linkUrlInput.value = 'https://';
    linkTextInput.value = '';
    linkModal.classList.add('open');
    setTimeout(() => linkUrlInput.focus(), 50);
  };
  document.getElementById('closeLinkModal')!.onclick = closeLinkModal;
  document.getElementById('cancelLinkBtn')!.onclick = closeLinkModal;
  linkModal.addEventListener('click', (e) => {
    if ((e.target as HTMLElement).id === 'linkModal') closeLinkModal();
  });
  linkInsertConfirmBtn.onclick = () => {
    const url = linkUrlInput.value.trim();
    const text = linkTextInput.value.trim() || url;
    if (!url) return;
    if (linkInsertMode === 'drive') {
      insertAtEditor(`<a href="${escapeAttr(url)}" target="_blank" rel="noopener">☁ ${escapeHtml(text)}</a>`);
      composeAttachments.push({ name: 'Google Drive', size: 0, type: 'drive', url });
      renderAttachments();
    } else {
      insertAtEditor(`<a href="${escapeAttr(url)}" target="_blank" rel="noopener">${escapeHtml(text)}</a>`);
    }
    closeLinkModal();
  };
  linkTextInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      document.getElementById('insertLinkConfirm')!.click();
    }
  });
  linkUrlInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      linkTextInput.focus();
    }
  });

  document.getElementById('attachFileBtn')!.onclick = () =>
    (document.getElementById('attachFileInput') as HTMLInputElement).click();
  document.getElementById('attachFileInput')!.addEventListener('change', (e) => {
    [...(e.target as HTMLInputElement).files!].forEach((f) =>
      composeAttachments.push({ name: f.name, size: f.size, type: f.type })
    );
    (e.target as HTMLInputElement).value = '';
    renderAttachments();
  });
  document.getElementById('insertImageBtn')!.onclick = () =>
    (document.getElementById('attachImageInput') as HTMLInputElement).click();
  document.getElementById('attachImageInput')!.addEventListener('change', (e) => {
    [...(e.target as HTMLInputElement).files!].forEach((f) => {
      const r = new FileReader();
      r.onload = (ev) =>
        insertAtEditor(`<img src="${ev.target?.result}" alt="${escapeAttr(f.name)}" />`);
      r.readAsDataURL(f);
    });
    (e.target as HTMLInputElement).value = '';
  });
  document.getElementById('insertDriveBtn')!.onclick = () => {
    linkInsertMode = 'drive';
    linkModalTitle.textContent = t('drive_url_title');
    linkInsertConfirmBtn.textContent = t('insert');
    linkUrlInput.value = 'https://drive.google.com/';
    linkTextInput.value = 'Google Drive';
    linkModal.classList.add('open');
    setTimeout(() => linkUrlInput.focus(), 50);
  };

  buildEmojiGrid();
  const emojiPop = document.getElementById('emojiPop')!;
  document.getElementById('insertEmojiBtn')!.addEventListener('click', (e) => {
    e.stopPropagation();
    const btn = (e.currentTarget as HTMLElement).getBoundingClientRect();
    const toolbar = document.getElementById('composeToolbar')!.getBoundingClientRect();
    emojiPop.style.left = btn.left - toolbar.left + 'px';
    emojiPop.style.top = btn.bottom - toolbar.top + 4 + 'px';
    emojiPop.classList.toggle('open');
  });
  document.addEventListener('click', (e) => {
    if (!emojiPop.contains(e.target as Node) && (e.target as HTMLElement).id !== 'insertEmojiBtn')
      emojiPop.classList.remove('open');
  });

  document.getElementById('insertSigBtn')!.onclick = () => {
    const sig = profile.signature || '';
    insertAtEditor('<br><br><div style="color:var(--fg-mute)">' + escapeHtml(sig).replace(/\n/g, '<br>') + '</div>');
  };

  const schedulePop = document.getElementById('schedulePop')!;
  document.getElementById('scheduleToggle')!.addEventListener('click', (e) => {
    e.stopPropagation();
    schedulePop.classList.toggle('open');
  });
  document.addEventListener('click', (e) => {
    if (!schedulePop.contains(e.target as Node) && (e.target as HTMLElement).id !== 'scheduleToggle')
      schedulePop.classList.remove('open');
  });
  document.querySelectorAll<HTMLButtonElement>('.schedule-quick button').forEach((b) => {
    b.onclick = () => {
      const now = new Date();
      const d = new Date(now);
      if (b.dataset.quick === '1h') d.setHours(now.getHours() + 1);
      else if (b.dataset.quick === 'tomorrow') {
        d.setDate(now.getDate() + 1);
        d.setHours(9, 0, 0, 0);
      } else if (b.dataset.quick === 'monday') {
        const off = (8 - now.getDay()) % 7 || 7;
        d.setDate(now.getDate() + off);
        d.setHours(9, 0, 0, 0);
      }
      scheduledFor = d.toISOString();
      updateComposeStatus();
      schedulePop.classList.remove('open');
    };
  });
  document.getElementById('scheduleConfirm')!.onclick = () => {
    const v = (document.getElementById('scheduleInput') as HTMLInputElement).value;
    if (!v) return;
    scheduledFor = new Date(v).toISOString();
    updateComposeStatus();
    schedulePop.classList.remove('open');
  };

  document.getElementById('closeReader')!.onclick = () =>
    document.getElementById('readerModal')!.classList.remove('open');
  document.getElementById('readerModal')!.addEventListener('click', (e) => {
    if ((e.target as HTMLElement).id === 'readerModal')
      (e.currentTarget as HTMLElement).classList.remove('open');
  });

  const userAvatar = document.getElementById('userAvatar')!;
  const userMenu = document.getElementById('userMenu')!;
  userAvatar.addEventListener('click', (e) => {
    e.stopPropagation();
    userMenu.classList.toggle('open');
  });
  document.addEventListener('click', (e) => {
    if (!userMenu.contains(e.target as Node) && e.target !== userAvatar) {
      userMenu.classList.remove('open');
    }
  });
  document.getElementById('logoutBtn')!.onclick = async () => {
    const confirmed = await showConfirm({
      title: t('logout'),
      message: t('confirm_logout'),
      confirmText: t('logout'),
      danger: true,
    });
    if (confirmed) {
      await supabase.auth.signOut();
      store.del(KEYS.session);
      window.location.href = './';
    }
  };

  const modal = document.getElementById('settingsModal')!;
  document.getElementById('openSettings')!.onclick = () => modal.classList.add('open');
  document.getElementById('closeSettings')!.onclick = () => modal.classList.remove('open');
  modal.addEventListener('click', (e) => {
    if (e.target === modal) modal.classList.remove('open');
  });

  document.querySelectorAll<HTMLElement>('.modal-nav .nav-item').forEach((n) => {
    n.onclick = () => {
      document.querySelectorAll('.modal-nav .nav-item').forEach((x) => x.classList.remove('active'));
      n.classList.add('active');
      document.querySelectorAll('.settings-section').forEach((s) => s.classList.remove('active'));
      document
        .querySelector<HTMLElement>(`.settings-section[data-section="${n.dataset.section}"]`)!
        .classList.add('active');
    };
  });



  const syncBtn = document.getElementById('syncInboxBtn');
  if (syncBtn) {
    syncBtn.onclick = async () => {
      let totalSynced = 0;
      if (activeAccount !== 'all') {
        const acc = accounts.find((a) => a.email === activeAccount);
        totalSynced = await syncAccountInbox(activeAccount, acc?.provider);
      } else {
        for (const a of accounts) {
          totalSynced += await syncAccountInbox(a.email, a.provider);
        }
      }
      renderAccounts();
      renderMails();
      toast(`${t('synced_success')} (${totalSynced} nuevos)`);
    };
  }

  // Sincronización automática periódica en tiempo real cada 30 segundos
  setInterval(async () => {
    if (accounts.length > 0) {
      let totalSynced = 0;
      for (const a of accounts) {
        const effToken = a.access_token || a.token || (window as any).supabaseProviderToken;
        if (effToken) {
          totalSynced += await syncAccountInbox(a.email, a.provider, effToken);
        }
      }
      if (totalSynced > 0) {
        renderAccounts();
        renderMails();
      }
    }
  }, 30000);

  const connectModal = document.getElementById('connectProviderModal');
  const closeConnect = document.getElementById('closeConnectModal');
  const skipConnect = document.getElementById('skipConnectModal');
  const googleBtn = document.getElementById('googleOAuthBtn');
  const msBtn = document.getElementById('microsoftOAuthBtn');
  const settingsGoogleBtn = document.getElementById('settingsGoogleOAuthBtn');
  const settingsMsBtn = document.getElementById('settingsMicrosoftOAuthBtn');

  const triggerGoogleOAuth = async () => {
    try {
      // 1. Usar supabase.auth.linkIdentity para vincular la identidad manteniendo la sesión actual de Supabase
      // Esto utiliza la URL de callback de Supabase (https://fbvgznyuzwpfcscxxfvr.supabase.co/auth/v1/callback) ya autorizada en Google Cloud
      const { error } = await supabase.auth.linkIdentity({
        provider: 'google',
        options: {
          scopes: GOOGLE_OAUTH_SCOPES,
          redirectTo: window.location.origin + window.location.pathname,
          queryParams: {
            prompt: 'consent',
            access_type: 'online',
          },
        },
      });

      if (error) {
        console.warn('Supabase linkIdentity error:', error.message);
        let clientId = (import.meta as any).env.PUBLIC_GOOGLE_CLIENT_ID || (window as any).PUBLIC_GOOGLE_CLIENT_ID || '1053272205885-03ecg2aig51gds1f1va4h55n4j5t1l2l.apps.googleusercontent.com';
        const redirectUri = encodeURIComponent(window.location.origin + window.location.pathname);
        const scope = encodeURIComponent(GOOGLE_OAUTH_SCOPES);
        window.location.href = `https://accounts.google.com/o/oauth2/v2/auth?client_id=${clientId.trim()}&redirect_uri=${redirectUri}&response_type=token&scope=${scope}&prompt=consent&access_type=online`;
      }
    } catch (err: any) {
      toast('⚠️ Error al conectar con Google OAuth.');
    }
  };

  const triggerMicrosoftOAuth = async () => {
    try {
      const { error } = await supabase.auth.linkIdentity({
        provider: 'azure',
        options: {
          scopes: MICROSOFT_OAUTH_SCOPES,
          redirectTo: window.location.origin + window.location.pathname,
          queryParams: { prompt: 'consent' },
        },
      });
      if (error) {
        const { error: signInErr } = await supabase.auth.signInWithOAuth({
          provider: 'azure',
          options: {
            scopes: MICROSOFT_OAUTH_SCOPES,
            redirectTo: window.location.origin + window.location.pathname,
            queryParams: { prompt: 'consent' },
          },
        });
        if (signInErr) {
          toast('⚠️ Microsoft OAuth no habilitado en Supabase.');
          console.warn('OAuth Microsoft Error:', signInErr.message);
        }
      }
    } catch (err: any) {
      toast('⚠️ Error al conectar con Microsoft OAuth.');
    }
  };

  if (googleBtn) googleBtn.onclick = triggerGoogleOAuth;
  if (settingsGoogleBtn) settingsGoogleBtn.onclick = triggerGoogleOAuth;
  if (msBtn) msBtn.onclick = triggerMicrosoftOAuth;
  if (settingsMsBtn) settingsMsBtn.onclick = triggerMicrosoftOAuth;

  const hideConnectModal = () => connectModal?.classList.remove('open');

  if (closeConnect) closeConnect.onclick = hideConnectModal;
  if (skipConnect) skipConnect.onclick = hideConnectModal;
  if (connectModal) {
    connectModal.addEventListener('click', (e) => {
      if (e.target === connectModal) hideConnectModal();
    });
  }

  const hasGmailOrOutlook = accounts.some((a) => a.provider === 'gmail' || a.provider === 'outlook' || a.email.includes('gmail') || a.email.includes('outlook'));
  if (!hasGmailOrOutlook && connectModal) {
    setTimeout(() => {
      connectModal.classList.add('open');
    }, 600);
  }

  const nameInp = document.getElementById('newCategoryName') as HTMLInputElement;
  const colorInp = document.getElementById('newCategoryColor') as HTMLInputElement;
  const addCat = async () => {
    const name = nameInp.value.trim();
    if (!name) return;
    if (categories.some((c) => c.name.toLowerCase() === name.toLowerCase())) {
      toast(`⚠️ La categoría "${name}" ya existe.`);
      return;
    }
    const color = colorInp.value || '#4c8bf5';
    nameInp.value = '';

    if (currentUser) {
      const { data: newDbCat, error } = await supabase.from('categories').insert({
        user_id: currentUser.id,
        name: name,
        color: color,
        keywords: [],
      }).select().single();

      if (error) {
        toast(`⚠️ Error al crear categoría en la base de datos: ${error.message}`);
        console.warn('Error al crear categoría:', error);
        return;
      }
      if (newDbCat) {
        categories.push({ id: newDbCat.id, name: newDbCat.name, color: newDbCat.color, keywords: newDbCat.keywords || [] });
      }
    } else {
      categories.push({ id: uid('c'), name, color, keywords: [] });
    }

    persist();
    renderCategoriesEditor();
    renderMails();
  };
  document.getElementById('addCategoryBtn')!.onclick = addCat;
  nameInp.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      addCat();
    }
  });
  nameInp.setAttribute('placeholder', t('category_name_placeholder'));

  document.getElementById('saveProfile')!.onclick = () => {
    profile = {
      name: (document.getElementById('profName') as HTMLInputElement).value,
      email: (document.getElementById('profEmail') as HTMLInputElement).value,
      phone: (document.getElementById('profPhone') as HTMLInputElement).value,
      signature: (document.getElementById('profSig') as HTMLTextAreaElement).value,
    };
    persist();
    hydrateProfileForm();
    const btn = document.getElementById('saveProfile')!;
    const orig = btn.textContent;
    btn.textContent = '✔';
    setTimeout(() => (btn.textContent = orig), 1200);
  };

  document.querySelectorAll<HTMLElement>('.theme-card').forEach((c) => {
    c.onclick = () => {
      store.set(KEYS.theme, c.dataset.theme!);
      applyTheme();
      hydrateAppearance();
    };
  });
  document.querySelectorAll<HTMLElement>('.swatch[data-color]').forEach((s) => {
    s.onclick = () => {
      store.set(KEYS.accent, s.dataset.color!);
      applyTheme();
      hydrateAppearance();
    };
  });
  document.getElementById('customColor')!.addEventListener('input', (e) => {
    store.set(KEYS.accent, (e.target as HTMLInputElement).value);
    applyTheme();
    hydrateAppearance();
  });

  document.querySelectorAll<HTMLElement>('.lang-btn').forEach((b) => {
    b.onclick = () => {
      store.set(KEYS.lang, b.dataset.lang!);
      applyI18n();
      nameInp.setAttribute('placeholder', t('category_name_placeholder'));
      renderAccounts();
      renderTabs();
      renderMails();
      renderCategoriesEditor();
      hydrateAppearance();
    };
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      const confirmModal = document.getElementById('confirmModal');
      if (confirmModal?.classList.contains('open')) {
        document.getElementById('confirmCancel')!.click();
        return;
      }
      document.getElementById('composeModal')!.classList.remove('open');
      document.getElementById('settingsModal')!.classList.remove('open');
      document.getElementById('readerModal')!.classList.remove('open');
      document.getElementById('linkModal')!.classList.remove('open');
      document.getElementById('emojiPop')!.classList.remove('open');
      document.getElementById('schedulePop')!.classList.remove('open');
      userMenu.classList.remove('open');
    }
  });

  // Desvincular cuentas de correo al cerrar el navegador (no en recarga de página)
  let skipCloseCleanup = false;
  window.addEventListener('keydown', (e) => {
    if (e.key === 'F5' || ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'r')) {
      skipCloseCleanup = true;
    }
  });
  window.addEventListener('beforeunload', () => {
    sessionStorage.setItem('triade.unloadTime', Date.now().toString());
  });
  window.addEventListener('pagehide', (event) => {
    if (skipCloseCleanup || event.persisted) return;
    unlinkAllAccountsOnBrowserClose();
  });
});
