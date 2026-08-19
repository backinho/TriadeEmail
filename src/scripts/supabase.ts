import { createClient } from '@supabase/supabase-js';

const url =
  (import.meta.env.PUBLIC_SUPABASE_URL as string) ||
  (import.meta.env.VITE_SUPABASE_URL as string) ||
  'https://fbvgznyuzwpfcscxxfvr.supabase.co';

const anonKey =
  (import.meta.env.PUBLIC_SUPABASE_ANON_KEY as string) ||
  (import.meta.env.VITE_SUPABASE_ANON_KEY as string) ||
  'sb_publishable_V8UYsBCqZJqEtjYV0xKeKg_9y9yOUQt';

export const supabaseUrl = url;
export const supabaseAnonKey = anonKey;

export const supabase = createClient(url, anonKey, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
    // Los tokens de las cuentas de correo llegan directamente a /app y no son sesiones de Triade.
    detectSessionInUrl: false,
  },
});

export default supabase;
