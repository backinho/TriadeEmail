-- =====================================================================
-- MIGRACIÓN 20260726000001: ACTUALIZACIÓN ESTRUCTURA USER_ACCOUNTS OAUTH
-- =====================================================================

-- Añadir columnas de tokens OAuth a user_accounts si no existen
ALTER TABLE public.user_accounts
ADD COLUMN IF NOT EXISTS access_token TEXT,
ADD COLUMN IF NOT EXISTS refresh_token TEXT,
ADD COLUMN IF NOT EXISTS expires_at BIGINT;

-- Modificar Foreign Key para apuntar directamente a auth.users(id)
DO $$
BEGIN
    IF EXISTS (
        SELECT 1 FROM information_schema.table_constraints 
        WHERE constraint_name = 'user_accounts_user_id_fkey' AND table_name = 'user_accounts'
    ) THEN
        ALTER TABLE public.user_accounts DROP CONSTRAINT user_accounts_user_id_fkey;
    END IF;
END $$;

ALTER TABLE public.user_accounts
ADD CONSTRAINT user_accounts_user_id_fkey 
FOREIGN KEY (user_id) REFERENCES auth.users(id) ON DELETE CASCADE;

-- Comentarios explicativos de la estructura
COMMENT ON TABLE public.user_accounts IS 'Cuentas de correo vinculadas al usuario con tokens OAuth de Google / Microsoft';
COMMENT ON COLUMN public.user_accounts.id IS 'Primary Key UUID';
COMMENT ON COLUMN public.user_accounts.user_id IS 'Foreign Key -> auth.users(id)';
COMMENT ON COLUMN public.user_accounts.email IS 'El correo de Gmail/Outlook conectado';
COMMENT ON COLUMN public.user_accounts.access_token IS 'Token de acceso de Google / Microsoft';
COMMENT ON COLUMN public.user_accounts.refresh_token IS 'Fundamental para no perder el acceso';
COMMENT ON COLUMN public.user_accounts.expires_at IS 'Cuándo vence el access token (timestamp Unix bigint)';
