import { createClient } from '@supabase/supabase-js';

const url =
  (import.meta.env.PUBLIC_SUPABASE_URL as string) ||
  (import.meta.env.VITE_SUPABASE_URL as string) ||
  'https://fbvgznyuzwpfcscxxfvr.supabase.co';

const anonKey =
  (import.meta.env.PUBLIC_SUPABASE_ANON_KEY as string) ||
  (import.meta.env.VITE_SUPABASE_ANON_KEY as string) ||
  'sb_publishable_V8UYsBCqZJqEtjYV0xKeKg_9y9yOUQt';

export const supabase = createClient(url, anonKey, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: true,
  },
});

export default supabase;
