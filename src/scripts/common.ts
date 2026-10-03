export type Lang = 'es' | 'en';
export type Theme = 'dark' | 'light';
export type EmailProvider = 'gmail' | 'outlook' | 'custom';

export interface Account {
  id?: string;
  user_id?: string;
  email: string;
  primary: boolean;
  provider?: EmailProvider;
  status?: 'connected' | 'syncing';
  lastSync?: string;
  access_token?: string;
  refresh_token?: string;
  expires_at?: number;
  token?: string;
}

export interface Category {
  id: string;
  name: string;
  color: string;
  keywords: string[];
}

export interface Attachment {
  name: string;
  size: number;
  type: string;
  url?: string;
  data?: string;
}

export type Folder = 'inbox' | 'sent' | 'drafts' | 'spam' | 'trash';

export interface Mail {
  id: string;
  from: string;
  fromEmail?: string;
  to: string;
  subject: string;
  body: string;
  bodyHtml?: string;
  account: string;
  time: string;
  timestamp?: number;
  unread: boolean;
  starred: boolean;
  folder: Folder;
  attachments?: Attachment[];
  scheduledFor?: string | null;
  draftDirty?: boolean;
}

export interface Profile {
  name: string;
  email: string;
  phone: string;
  signature: string;
}

export const KEYS = {
  theme: 'triade.theme',
  accent: 'triade.accent',
  lang: 'triade.lang',
  accounts: 'triade.accounts',
  categories: 'triade.categories',
  profile: 'triade.profile',
  pageSize: 'triade.pageSize',
  mails: 'triade.mails',
  session: 'triade.session',
  sessionAccounts: 'triade.sessionAccounts',
  unlinked: 'triade.unlinked',
  unlinkNotice: 'triade.unlinkNotice',
} as const;

export type StorageKey = (typeof KEYS)[keyof typeof KEYS];

export const store = {
  get<T>(k: StorageKey | string, def: T): T {
    try {
      const v = localStorage.getItem(k);
      return v ? (JSON.parse(v) as T) : def;
    } catch {
      return def;
    }
  },
  set(k: StorageKey | string, v: unknown): void {
    try {
      localStorage.setItem(k, JSON.stringify(v));
    } catch (err) {
      console.warn(`[Storage] No se pudo guardar '${k}' en localStorage (cuota excedida o almacenamiento bloqueado):`, err);
    }
  },
  del(k: StorageKey | string): void {
    try {
      localStorage.removeItem(k);
    } catch {
      // Ignore
    }
  },
};

export const I18N: Record<Lang, Record<string, string>> = {
  es: {
    welcome_back: 'Bienvenido de nuevo',
    sign_in_sub: 'Accede a tu bandeja unificada Triade',
    email: 'Correo electrónico',
    password: 'Contraseña',
    remember: 'Recordarme',
    forgot: '¿Olvidaste tu contraseña?',
    sign_in: 'Iniciar sesión',
    no_account: '¿No tienes cuenta?',
    create_account: 'Crear cuenta',
    create_title: 'Crea tu cuenta',
    create_sub: 'Une todos tus correos en un solo lugar',
    full_name: 'Nombre completo',
    confirm_password: 'Confirmar contraseña',
    already_account: '¿Ya tienes cuenta?',
    sign_in_link: 'Inicia sesión',
    recover_title: 'Recuperar contraseña',
    recover_sub: 'Te enviaremos un enlace de recuperación',
    send_link: 'Enviar enlace',
    back_to_login: 'Volver al inicio de sesión',
    compose: 'Redactar',
    search_placeholder: 'Buscar correos, remitentes, palabras clave…',
    inbox: 'Bandeja de entrada',
    starred: 'Destacados',
    sent: 'Enviados',
    drafts: 'Borradores',
    spam: 'Spam',
    trash: 'Papelera',
    accounts: 'Cuentas',
    settings: 'Ajustes',
    all: 'Todos',
    uncategorized: 'Sin categoría',
    parameters: 'Parámetros',
    profile: 'Perfil',
    appearance: 'Apariencia e idioma',
    manage_accounts: 'Cuentas de correo',
    manage_accounts_hint:
      'Añade y gestiona los correos que centralizarás en Triade.',
    add_account: 'Añadir correo',
    keywords_title: 'Categorías y palabras clave',
    keywords_hint:
      'Crea tus propias categorías. Cada categoría se convierte en una pestaña que filtra los correos cuyo asunto contenga alguna de sus palabras clave.',
    add_category: 'Añadir categoría',
    category_name_placeholder: 'Nombre de la categoría…',
    new_keyword_placeholder: 'Nueva palabra clave…',
    add: 'Añadir',
    profile_hint: 'Datos que se muestran en tu cuenta.',
    display_name: 'Nombre a mostrar',
    phone: 'Teléfono',
    signature: 'Firma',
    save_changes: 'Guardar cambios',
    theme: 'Tema',
    theme_dark: 'Oscuro',
    theme_light: 'Claro',
    accent_color: 'Color de acento',
    pick_custom: 'Personalizado',
    language: 'Idioma',
    spanish: 'Español',
    english: 'Inglés',
    primary_account: 'Principal',
    remove: 'Quitar',
    no_mails: 'No hay correos que coincidan con este filtro.',
    no_categories: 'Aún no has creado categorías. Añade una para empezar a filtrar.',
    per_page: 'por página',
    showing: 'Mostrando',
    of: 'de',
    prev: 'Anterior',
    next: 'Siguiente',
    compose_title: 'Nuevo mensaje',
    to: 'Para',
    cc: 'CC',
    subject: 'Asunto',
    body: 'Mensaje',
    send: 'Enviar',
    save_draft: 'Guardar borrador',
    discard: 'Descartar',
    compose_placeholder: 'Escribe tu mensaje…',
    subject_placeholder: 'Asunto del correo',
    to_placeholder: 'destinatario@correo.com',
    star: 'Destacar',
    unstar: 'Quitar destacado',
    move_spam: 'Marcar spam',
    move_trash: 'Enviar a papelera',
    restore: 'Restaurar a bandeja',
    mail_sent: 'Correo enviado',
    draft_saved: 'Borrador guardado',
    logout: 'Cerrar sesión',
    signed_in_as: 'Sesión iniciada como',
    confirm_logout: '¿Cerrar sesión de Triade Mail?',
    confirm_unlink: '¿Estás seguro de que deseas desvincular la cuenta {email}?',
    confirm_relink: '¿Deseas vincular la cuenta ahora para obtener los permisos de envío?',
    confirm: 'Confirmar',
    token_prompt_title: 'Token OAuth',
    token_prompt_message: 'Ingresa tu Token OAuth de Google/Microsoft (Bearer Key ya29...) para {email}:',
    drive_url_title: 'Enlace de Google Drive',
    empty_inbox: 'Tu bandeja está vacía.',
    empty_starred: 'No tienes correos destacados.',
    empty_sent: 'Aún no has enviado correos.',
    empty_drafts: 'No tienes borradores guardados.',
    empty_spam: 'No hay correos en spam.',
    empty_trash: 'La papelera está vacía.',
    menu: 'Menú',
    schedule_send: 'Programar envío',
    schedule: 'Programar',
    scheduled: 'Programado',
    scheduled_for: 'Programado para',
    reply: 'Responder',
    forward: 'Reenviar',
    delete_forever: 'Eliminar definitivamente',
    confirm_delete: '¿Eliminar este correo definitivamente?',
    mail_deleted: 'Correo eliminado',
    mail_scheduled: 'Correo programado',
    link_url_prompt: 'URL del enlace:',
    link_text_prompt: 'Texto a mostrar (opcional):',
    drive_url_prompt: 'Enlace de Google Drive:',
    attachment: 'Adjunto',
    insert_link: 'Insertar enlace',
    link_url: 'URL del enlace',
    link_text: 'Texto a mostrar',
    cancel: 'Cancelar',
    insert: 'Insertar',
    sync_inbox: 'Sincronizar bandeja',
    syncing: 'Sincronizando...',
    synced_success: 'Bandeja sincronizada correctamente',
    connect_provider_title: 'Registra tu cuenta de correo',
    connect_provider_desc: 'Conecta tu cuenta de Outlook para centralizar tus mensajes en Triade.',
    select_provider: 'Proveedor',
    connect_sync: 'Conectar y Sincronizar',
    gmail: 'Gmail',
    outlook: 'Outlook',
    custom_provider: 'Otro / Personalizado',
    last_synced: 'Última sincr.:',
    from_account: 'De',
    select_account: 'Seleccionar cuenta…',
    no_linked_accounts: 'No hay cuentas vinculadas para enviar',
    send_error_no_account: 'Selecciona una cuenta vinculada para enviar',
    send_error_no_token: 'Token expirado o no disponible. Vuelve a vincular la cuenta.',
    send_error_failed: 'Error al enviar el correo. Verifica la cuenta y los permisos.',
    account_unlinked: 'Cuenta {email} desvinculada exitosamente.',
    accounts_unlinked_on_close: 'Las cuentas vinculadas fueron desvinculadas al cerrar el navegador: {emails}',
    account_unlinked_on_close: 'La cuenta {email} fue desvinculada al cerrar el navegador.',
  },
  en: {
    welcome_back: 'Welcome back',
    sign_in_sub: 'Access your Triade unified inbox',
    email: 'Email',
    password: 'Password',
    remember: 'Remember me',
    forgot: 'Forgot your password?',
    sign_in: 'Sign in',
    no_account: "Don't have an account?",
    create_account: 'Create account',
    create_title: 'Create your account',
    create_sub: 'Bring all your emails into one place',
    full_name: 'Full name',
    confirm_password: 'Confirm password',
    already_account: 'Already have an account?',
    sign_in_link: 'Sign in',
    recover_title: 'Recover password',
    recover_sub: "We'll send you a recovery link",
    send_link: 'Send link',
    back_to_login: 'Back to sign in',
    compose: 'Compose',
    search_placeholder: 'Search mail, senders, keywords…',
    inbox: 'Inbox',
    starred: 'Starred',
    sent: 'Sent',
    drafts: 'Drafts',
    spam: 'Spam',
    trash: 'Trash',
    accounts: 'Accounts',
    settings: 'Settings',
    all: 'All',
    uncategorized: 'Uncategorized',
    parameters: 'Parameters',
    profile: 'Profile',
    appearance: 'Appearance & language',
    manage_accounts: 'Email accounts',
    manage_accounts_hint:
      'Add and manage the emails you centralize in Triade.',
    add_account: 'Add email',
    keywords_title: 'Categories & keywords',
    keywords_hint:
      'Create your own categories. Each becomes a tab that filters mail whose subject contains any of its keywords.',
    add_category: 'Add category',
    category_name_placeholder: 'Category name…',
    new_keyword_placeholder: 'New keyword…',
    add: 'Add',
    profile_hint: 'Details shown on your account.',
    display_name: 'Display name',
    phone: 'Phone',
    signature: 'Signature',
    save_changes: 'Save changes',
    theme: 'Theme',
    theme_dark: 'Dark',
    theme_light: 'Light',
    accent_color: 'Accent color',
    pick_custom: 'Custom',
    language: 'Language',
    spanish: 'Spanish',
    english: 'English',
    primary_account: 'Primary',
    remove: 'Remove',
    no_mails: 'No mail matches this filter.',
    no_categories: "You haven't created any categories yet. Add one to start filtering.",
    per_page: 'per page',
    showing: 'Showing',
    of: 'of',
    prev: 'Previous',
    next: 'Next',
    compose_title: 'New message',
    to: 'To',
    cc: 'CC',
    subject: 'Subject',
    body: 'Message',
    send: 'Send',
    save_draft: 'Save draft',
    discard: 'Discard',
    compose_placeholder: 'Write your message…',
    subject_placeholder: 'Email subject',
    to_placeholder: 'recipient@mail.com',
    star: 'Star',
    unstar: 'Unstar',
    move_spam: 'Mark as spam',
    move_trash: 'Move to trash',
    restore: 'Restore to inbox',
    mail_sent: 'Message sent',
    draft_saved: 'Draft saved',
    logout: 'Sign out',
    signed_in_as: 'Signed in as',
    confirm_logout: 'Sign out of Triade Mail?',
    confirm_unlink: 'Are you sure you want to unlink account {email}?',
    confirm_relink: 'Do you want to link the account now to get send permissions?',
    confirm: 'Confirm',
    token_prompt_title: 'OAuth Token',
    token_prompt_message: 'Enter your Google/Microsoft OAuth token (Bearer Key ya29...) for {email}:',
    drive_url_title: 'Google Drive link',
    empty_inbox: 'Your inbox is empty.',
    empty_starred: 'No starred messages.',
    empty_sent: "You haven't sent any messages yet.",
    empty_drafts: 'No saved drafts.',
    empty_spam: 'No spam messages.',
    empty_trash: 'Trash is empty.',
    menu: 'Menu',
    schedule_send: 'Schedule send',
    schedule: 'Schedule',
    scheduled: 'Scheduled',
    scheduled_for: 'Scheduled for',
    reply: 'Reply',
    forward: 'Forward',
    delete_forever: 'Delete forever',
    confirm_delete: 'Permanently delete this message?',
    mail_deleted: 'Message deleted',
    mail_scheduled: 'Message scheduled',
    link_url_prompt: 'Link URL:',
    link_text_prompt: 'Text to display (optional):',
    drive_url_prompt: 'Google Drive link:',
    attachment: 'Attachment',
    insert_link: 'Insert link',
    link_url: 'Link URL',
    link_text: 'Display text',
    cancel: 'Cancel',
    insert: 'Insert',
    sync_inbox: 'Sync inbox',
    syncing: 'Syncing...',
    synced_success: 'Inbox synchronized successfully',
    connect_provider_title: 'Register your email account',
    connect_provider_desc: 'Connect your Outlook account to centralize your messages in Triade.',
    select_provider: 'Provider',
    connect_sync: 'Connect & Sync',
    gmail: 'Gmail',
    outlook: 'Outlook',
    custom_provider: 'Other / Custom',
    last_synced: 'Last synced:',
    from_account: 'From',
    select_account: 'Select account…',
    no_linked_accounts: 'No linked accounts available to send',
    send_error_no_account: 'Select a linked account to send',
    send_error_no_token: 'Token expired or unavailable. Re-link the account.',
    send_error_failed: 'Failed to send email. Check account and permissions.',
    account_unlinked: 'Account {email} unlinked successfully.',
    accounts_unlinked_on_close: 'Linked accounts were unlinked when the browser closed: {emails}',
    account_unlinked_on_close: 'Account {email} was unlinked when the browser closed.',
  },
};

export function applyTheme(): void {
  const theme = store.get<Theme>(KEYS.theme, 'light');
  const accent = store.get<string>(KEYS.accent, '#e63946');
  document.documentElement.setAttribute('data-theme', theme);
  document.documentElement.style.setProperty('--accent', accent);
  const c = accent.replace('#', '');
  const r = parseInt(c.substring(0, 2), 16);
  const g = parseInt(c.substring(2, 4), 16);
  const b = parseInt(c.substring(4, 6), 16);
  document.documentElement.style.setProperty('--accent-glow', `${r}, ${g}, ${b}`);
}

export function t(key: string): string {
  const lang = store.get<Lang>(KEYS.lang, 'es');
  return (I18N[lang] && I18N[lang][key]) || I18N.es[key] || key;
}

export function applyI18n(root: Document | HTMLElement = document): void {
  root.querySelectorAll('[data-i18n]').forEach((el) => {
    el.textContent = t((el as HTMLElement).dataset.i18n!);
  });
  root.querySelectorAll('[data-i18n-placeholder]').forEach((el) => {
    (el as HTMLElement).setAttribute(
      'placeholder',
      t((el as HTMLElement).dataset.i18nPlaceholder!)
    );
  });
  root.querySelectorAll('[data-i18n-title]').forEach((el) => {
    (el as HTMLElement).setAttribute(
      'title',
      t((el as HTMLElement).dataset.i18nTitle!)
    );
  });
}

declare global {
  interface Window {
    Triade?: typeof import('./common');
  }
}

applyTheme();
document.addEventListener('DOMContentLoaded', () => applyI18n());
