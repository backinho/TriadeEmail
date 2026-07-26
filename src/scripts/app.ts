import {
  KEYS, store, t, applyTheme, applyI18n,
  type Account, type Category, type Profile, type Mail, type Attachment, type Folder,
} from './common';
import { supabase } from './supabase';


// ---------- Defaults ----------
const DEFAULT_ACCOUNTS: Account[] = [
  { email: 'operaciones@triade.com', primary: true },
  { email: 'ventas@triade.com', primary: false },
];
const DEFAULT_CATEGORIES: Category[] = [
  { id: 'c_work', name: 'Trabajo', color: '#e63946', keywords: ['reunión', 'meeting', 'proyecto', 'project', 'informe', 'report'] },
  { id: 'c_finance', name: 'Finanzas', color: '#4c8bf5', keywords: ['factura', 'invoice', 'pago', 'payment', 'presupuesto', 'budget'] },
  { id: 'c_promo', name: 'Promociones', color: '#a06eff', keywords: ['oferta', 'offer', 'descuento', 'discount', 'promo'] },
  { id: 'c_social', name: 'Social', color: '#3ccf91', keywords: ['invitación', 'invitation', 'evento', 'event', 'conexión'] },
];
const DEFAULT_PROFILE: Profile = {
  name: 'Operador Triade',
  email: 'operaciones@triade.com',
  phone: '+58 000 000 0000',
  signature: '— Triade · Levantamiento Artificial y Rehabilitación de Pozos',
};

const SUBJECTS: [string, string, string][] = [
  ['Halliburton', 'Factura Nº 4192 — VFD 200HP', 'Adjuntamos la factura del servicio de instalación…'],
  ['PDVSA Occidente', 'Reunión de coordinación proyecto Lago', 'Confirmamos reunión para revisar cronograma de rehabilitación…'],
  ['Schlumberger', 'Presupuesto BCP — 4 unidades', 'Enviamos el presupuesto solicitado para los equipos BCP…'],
  ['Quick Connectors Inc.', 'Update sistema P-5000', 'Actualización de firmware disponible para conectores P-5000…'],
  ['LinkedIn', 'Nueva conexión: Ing. Carla Méndez', 'Tienes una nueva invitación de conexión en tu red…'],
  ['Amazon Business', 'Oferta 20% en herramientas industriales', 'Descuento especial para clientes empresariales…'],
  ['RRHH Triade', 'Informe mensual de operaciones', 'Adjunto el reporte consolidado del mes…'],
  ['Weatherford', 'Newsletter — Boletín técnico Q3', 'Novedades trimestrales en cabezales y sensores…'],
  ['Eventos Petroleros', 'Invitación al evento OTC 2026', 'Tenemos el gusto de invitarle al evento anual…'],
  ['CANTV', 'Pago procesado correctamente', 'Su pago ha sido recibido y procesado…'],
  ['Baker Hughes', 'Cotización cabezales 150HP', 'Envío cotización actualizada para cabezales…'],
  ['SAP Concur', 'Reporte de gastos aprobado', 'Su reporte de gastos ha sido aprobado por el supervisor…'],
  ['Google Workspace', 'Actualización de seguridad', 'Se ha detectado un inicio de sesión desde un nuevo dispositivo…'],
  ['Cámara Petrolera', 'Invitación foro energético 2026', 'Le invitamos al próximo foro sectorial…'],
  ['DHL Express', 'Envío entregado — guía 8842', 'Su paquete ha sido entregado exitosamente…'],
  ['Banco Provincial', 'Factura de servicios corporativos', 'Adjuntamos su factura mensual…'],
  ['Microsoft 365', 'Reunión programada: Kickoff proyecto', 'Recordatorio: reunión mañana a las 10:00…'],
  ['Slack', 'Nuevo mensaje en #operaciones', 'Tienes 5 mensajes sin leer en el canal…'],
  ['Zoom', 'Grabación de reunión disponible', 'La grabación de la reunión de ayer está lista…'],
  ['Dropbox', 'Un archivo fue compartido contigo', 'El archivo "Cronograma_Q4.pdf" fue compartido…'],
  ['GitHub', 'Pull request abierto: firmware v2.3', 'Nueva PR pendiente de revisión…'],
  ['Adobe Creative', 'Renovación de suscripción', 'Su suscripción se renovará próximamente…'],
  ['Netflix', 'Descuento en tu próximo mes', 'Oferta especial disponible por tiempo limitado…'],
  ['Uber', 'Recibo de tu viaje', 'Gracias por viajar con nosotros. Adjunto recibo…'],
  ['Booking.com', 'Confirmación de reserva', 'Su reserva ha sido confirmada exitosamente…'],
  ['SENIAT', 'Notificación tributaria', 'Le informamos sobre su declaración pendiente…'],
  ['IEEE', 'Newsletter mensual', 'Últimas noticias del sector energético y automatización…'],
  ['Trello', 'Actualización en tablero Operaciones', 'Se agregaron 3 nuevas tarjetas al tablero…'],
  ['Notion', 'Documento actualizado por el equipo', 'El documento "Procedimientos" fue editado…'],
  ['AWS', 'Factura mensual de servicios', 'Su factura de servicios en la nube está disponible…'],
];

function timeLabel(i: number): string {
  if (i < 3) return `${9 + i}:${String((i * 13) % 60).padStart(2, '0')}`;
  if (i < 7) return 'Ayer';
  if (i < 12) return ['Lun', 'Mar', 'Mié', 'Jue', 'Vie'][i - 7];
  const day = 28 - (i - 12);
  return `${day} Sep`;
}

const uid = (p = 'm'): string =>
  p + '_' + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-3);

function seedMails(): Mail[] {
  return SUBJECTS.map((s, i) => ({
    id: uid('m'),
    from: s[0],
    to: '',
    subject: s[1],
    body: s[2],
    account: i % 3 === 0 ? 'ventas@triade.com' : 'operaciones@triade.com',
    time: timeLabel(i),
    unread: i < 6,
    starred: [1, 8, 14].includes(i),
    folder: 'inbox' as Folder,
  }));
}

// ---------- State ----------
let accounts: Account[] = store.get<Account[]>(KEYS.accounts, DEFAULT_ACCOUNTS);
let categories: Category[] = store.get<Category[]>(KEYS.categories, DEFAULT_CATEGORIES);
let profile: Profile = store.get<Profile>(KEYS.profile, DEFAULT_PROFILE);
let pageSize: number = store.get<number>(KEYS.pageSize, 10);
let mails: Mail[] = store.get<Mail[]>(KEYS.mails, []);
if (!Array.isArray(mails) || !mails.length) {
  mails = seedMails();
  store.set(KEYS.mails, mails);
}

let activeFolder: string = 'inbox';
let activeAccount: string = 'all';
let activeTab: string = 'all';
let searchQuery = '';
let currentPage = 1;

let currentUser: any = null;

const persist = (): void => {
  store.set(KEYS.accounts, accounts);
  store.set(KEYS.categories, categories);
  store.set(KEYS.profile, profile);
  store.set(KEYS.pageSize, pageSize);
  store.set(KEYS.mails, mails);

  // Sync profile & preferences to Supabase asynchronously
  if (currentUser) {
    supabase.from('profiles').upsert({
      id: currentUser.id,
      name: profile.name,
      email: profile.email,
      phone: profile.phone,
      signature: profile.signature,
      page_size: pageSize,
    }).then(({ error }) => {
      if (error) console.warn('Error al actualizar perfil en Supabase:', error);
    });
  }
};

async function syncSupabaseData(): Promise<void> {
  try {
    const { data: { session } } = await supabase.auth.getSession();
    if (!session || !session.user) {
      window.location.href = './';
      return;
    }
    currentUser = session.user;

    // 1. Cargar o inicializar Perfil
    const { data: pData } = await supabase.from('profiles').select('*').eq('id', currentUser.id).single();
    if (pData) {
      profile = {
        name: pData.name || currentUser.user_metadata?.full_name || 'Operador Triade',
        email: pData.email || currentUser.email || 'operaciones@triade.com',
        phone: pData.phone || '+58 000 000 0000',
        signature: pData.signature || '— Triade · Levantamiento Artificial y Rehabilitación de Pozos',
      };
    } else {
      profile.email = currentUser.email || profile.email;
      profile.name = currentUser.user_metadata?.full_name || profile.name;
    }

    // 2. Cargar Cuentas del usuario
    const { data: aData } = await supabase.from('user_accounts').select('*').eq('user_id', currentUser.id);
    if (aData && aData.length > 0) {
      accounts = aData.map((a: any) => ({ email: a.email, primary: a.is_primary }));
    }

    // 3. Cargar Categorías
    const { data: cData } = await supabase.from('categories').select('*').eq('user_id', currentUser.id);
    if (cData && cData.length > 0) {
      categories = cData.map((c: any) => ({ id: c.id, name: c.name, color: c.color, keywords: c.keywords || [] }));
    }

    // 4. Cargar Correos
    const { data: mData } = await supabase.from('mails').select('*').eq('user_id', currentUser.id).order('created_at', { ascending: false });
    if (mData && mData.length > 0) {
      mails = mData.map((m: any) => ({
        id: m.id,
        from: m.from_name,
        fromEmail: m.from_email || '',
        to: m.to_address,
        subject: m.subject,
        body: m.body,
        bodyHtml: m.body_html || '',
        account: m.account,
        time: m.time_label || new Date(m.created_at).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
        unread: m.unread,
        starred: m.starred,
        folder: m.folder as Folder,
        scheduledFor: m.scheduled_for,
      }));
    } else {
      // Seed inicial en Supabase si es la primera vez que ingresa
      const seeded = seedMails();
      mails = seeded;
      const dbRows = seeded.map((m) => ({
        user_id: currentUser.id,
        from_name: m.from,
        to_address: m.to || m.account,
        subject: m.subject,
        body: m.body,
        account: m.account,
        folder: m.folder,
        unread: m.unread,
        starred: m.starred,
        time_label: m.time,
      }));
      await supabase.from('mails').insert(dbRows);
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
  if (activeTab === 'all') return base;
  if (activeTab === '__uncat') return base.filter((m) => classify(m.subject) === '__uncat');
  return base.filter((m) => classify(m.subject) === activeTab);
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
    el.innerHTML = `<span class="dot"></span><span>${escapeHtml(a.email)}</span>`;
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
      row.innerHTML = `
        <span class="email">${escapeHtml(a.email)}</span>
        ${a.primary ? `<span class="badge">${escapeHtml(t('primary_account'))}</span>` : ''}
        <button class="btn sm ghost" data-remove="${i}">${escapeHtml(t('remove'))}</button>`;
      edit.appendChild(row);
    });
    edit.querySelectorAll<HTMLElement>('[data-remove]').forEach((b) => {
      b.onclick = () => {
        accounts.splice(+b.dataset.remove!, 1);
        persist();
        renderAccounts();
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
    if (!confirm(t('confirm_delete'))) return;
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
      categories.splice(idx, 1);
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
function toast(msg: string): void {
  const el = document.getElementById('toast')!;
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), 2200);
}

// ---------- Compose ----------
let composingDraftId: string | null = null;
let composeAttachments: Attachment[] = [];
let scheduledFor: string | null = null;

type DraftLike = Partial<Mail> & { id?: string };

function openCompose(draft?: DraftLike): void {
  const modal = document.getElementById('composeModal')!;
  composingDraftId = draft?.id ?? null;
  composeAttachments = draft?.attachments ? [...draft.attachments] : [];
  scheduledFor = draft?.scheduledFor ?? null;
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

function sendMail(): void {
  const f = readComposeFields();
  if (!f.to) {
    document.getElementById('composeTo')!.focus();
    return;
  }
  if (composingDraftId) {
    const idx = mails.findIndex((x) => x.id === composingDraftId);
    if (idx >= 0) mails.splice(idx, 1);
  }
  const now = new Date();
  const timeStr = scheduledFor
    ? new Date(scheduledFor).toLocaleString([], { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })
    : `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
  const acct = accounts.find((a) => a.primary)?.email || accounts[0]?.email || profile.email;
  mails.unshift({
    id: uid('m'),
    from: profile.name || acct,
    fromEmail: acct,
    to: f.to,
    subject: f.subject || '(sin asunto)',
    body: (f.body || '').split('\n')[0].slice(0, 180),
    bodyHtml: f.bodyHtml,
    attachments: composeAttachments.slice(),
    account: acct,
    time: timeStr,
    unread: false,
    starred: false,
    folder: 'sent',
    scheduledFor: scheduledFor || null,
  });
  persist();
  toast(scheduledFor ? t('mail_scheduled') : t('mail_sent'));
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
  const acct = accounts.find((a) => a.primary)?.email || accounts[0]?.email || profile.email;
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
document.addEventListener('DOMContentLoaded', () => {
  applyI18n();
  renderAccounts();
  renderMails();
  renderCategoriesEditor();
  hydrateProfileForm();
  hydrateAppearance();
  syncSupabaseData();


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
  const closeLinkModal = () => linkModal.classList.remove('open');

  document.getElementById('insertLinkBtn')!.onclick = () => {
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
  document.getElementById('insertLinkConfirm')!.onclick = () => {
    const url = linkUrlInput.value.trim();
    const text = linkTextInput.value.trim() || url;
    if (!url) return;
    insertAtEditor(`<a href="${escapeAttr(url)}" target="_blank" rel="noopener">${escapeHtml(text)}</a>`);
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
    const url = prompt(t('drive_url_prompt'), 'https://drive.google.com/');
    if (!url) return;
    insertAtEditor(`<a href="${escapeAttr(url)}" target="_blank" rel="noopener">☁ ${escapeHtml(url)}</a>`);
    composeAttachments.push({ name: 'Google Drive', size: 0, type: 'drive', url });
    renderAttachments();
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
    if (confirm(t('confirm_logout'))) {
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

  document.getElementById('addAccountBtn')!.onclick = () => {
    const inp = document.getElementById('newAccountEmail') as HTMLInputElement;
    const v = inp.value.trim();
    if (!v || !v.includes('@')) return;
    accounts.push({ email: v, primary: accounts.length === 0 });
    inp.value = '';
    persist();
    renderAccounts();
  };

  const nameInp = document.getElementById('newCategoryName') as HTMLInputElement;
  const colorInp = document.getElementById('newCategoryColor') as HTMLInputElement;
  const addCat = () => {
    const name = nameInp.value.trim();
    if (!name) return;
    categories.push({ id: uid('c'), name, color: colorInp.value || '#4c8bf5', keywords: [] });
    nameInp.value = '';
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
      document.getElementById('composeModal')!.classList.remove('open');
      document.getElementById('settingsModal')!.classList.remove('open');
      document.getElementById('readerModal')!.classList.remove('open');
      document.getElementById('linkModal')!.classList.remove('open');
      document.getElementById('emojiPop')!.classList.remove('open');
      document.getElementById('schedulePop')!.classList.remove('open');
      userMenu.classList.remove('open');
    }
  });
});
