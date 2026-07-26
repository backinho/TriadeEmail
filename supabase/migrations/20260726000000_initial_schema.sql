-- =====================================================================
-- MIGRACIÓN INICIAL SUPABASE: TRIADE MAIL
-- =====================================================================

-- ---------------------------------------------------------------------
-- 1. TABLA DE PERFILES (profiles)
-- ---------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.profiles (
    id UUID PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
    name TEXT NOT NULL DEFAULT '',
    email TEXT NOT NULL,
    phone TEXT DEFAULT '+58 000 000 0000',
    signature TEXT DEFAULT '— Triade · Levantamiento Artificial y Rehabilitación de Pozos',
    page_size INTEGER DEFAULT 10 CHECK (page_size > 0),
    theme TEXT DEFAULT 'light' CHECK (theme IN ('light', 'dark')),
    accent_color TEXT DEFAULT '#e63946',
    lang TEXT DEFAULT 'es' CHECK (lang IN ('es', 'en')),
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

COMMENT ON TABLE public.profiles IS 'Perfil del usuario y sus preferencias de aplicación';

-- ---------------------------------------------------------------------
-- 2. TABLA DE CUENTAS DE CORREO (user_accounts)
-- ---------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.user_accounts (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    user_id UUID NOT NULL REFERENCES public.profiles(id) ON DELETE CASCADE,
    email TEXT NOT NULL,
    is_primary BOOLEAN DEFAULT FALSE NOT NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
    UNIQUE(user_id, email)
);

COMMENT ON TABLE public.user_accounts IS 'Cuentas de correo vinculadas al usuario centralizado';

-- ---------------------------------------------------------------------
-- 3. TABLA DE CATEGORÍAS (categories)
-- ---------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.categories (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    user_id UUID NOT NULL REFERENCES public.profiles(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    color TEXT DEFAULT '#e63946' NOT NULL,
    keywords TEXT[] DEFAULT '{}'::TEXT[] NOT NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

COMMENT ON TABLE public.categories IS 'Categorías personalizadas y sus palabras clave para filtrado inteligente';

-- ---------------------------------------------------------------------
-- 4. TABLA DE CORREOS (mails)
-- ---------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.mails (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    user_id UUID NOT NULL REFERENCES public.profiles(id) ON DELETE CASCADE,
    from_name TEXT NOT NULL,
    from_email TEXT,
    to_address TEXT NOT NULL,
    subject TEXT DEFAULT '' NOT NULL,
    body TEXT DEFAULT '' NOT NULL,
    body_html TEXT,
    account TEXT NOT NULL,
    folder TEXT DEFAULT 'inbox' NOT NULL CHECK (folder IN ('inbox', 'sent', 'drafts', 'spam', 'trash')),
    unread BOOLEAN DEFAULT TRUE NOT NULL,
    starred BOOLEAN DEFAULT FALSE NOT NULL,
    scheduled_for TIMESTAMP WITH TIME ZONE,
    time_label TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

COMMENT ON TABLE public.mails IS 'Bandeja de mensajes entrantes, salientes, borradores y destacados';

-- ---------------------------------------------------------------------
-- 5. TABLA DE ADJUNTOS (attachments)
-- ---------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.attachments (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    mail_id UUID NOT NULL REFERENCES public.mails(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    size BIGINT NOT NULL,
    type TEXT NOT NULL,
    url TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

COMMENT ON TABLE public.attachments IS 'Archivos adjuntos vinculados a un correo electrónico';

-- ---------------------------------------------------------------------
-- 6. ÍNDICES DE RENDIMIENTO
-- ---------------------------------------------------------------------
CREATE INDEX IF NOT EXISTS idx_mails_user_folder ON public.mails(user_id, folder);
CREATE INDEX IF NOT EXISTS idx_mails_user_account ON public.mails(user_id, account);
CREATE INDEX IF NOT EXISTS idx_attachments_mail_id ON public.attachments(mail_id);
CREATE INDEX IF NOT EXISTS idx_categories_user_id ON public.categories(user_id);
CREATE INDEX IF NOT EXISTS idx_user_accounts_user_id ON public.user_accounts(user_id);

-- ---------------------------------------------------------------------
-- 7. FUNCIONES Y TRIGGERS (updated_at)
-- ---------------------------------------------------------------------
CREATE OR REPLACE FUNCTION public.handle_updated_at()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = timezone('utc'::text, now());
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER set_profiles_updated_at
    BEFORE UPDATE ON public.profiles
    FOR EACH ROW
    EXECUTE FUNCTION public.handle_updated_at();

CREATE TRIGGER set_mails_updated_at
    BEFORE UPDATE ON public.mails
    FOR EACH ROW
    EXECUTE FUNCTION public.handle_updated_at();

-- ---------------------------------------------------------------------
-- 8. TRIGGER AUTOMÁTICO DE NUEVO USUARIO (auth.users -> public.profiles)
-- ---------------------------------------------------------------------
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS TRIGGER AS $$
DECLARE
    user_email TEXT;
    user_full_name TEXT;
BEGIN
    user_email := NEW.email;
    user_full_name := COALESCE(NEW.raw_user_meta_data->>'full_name', SPLIT_PART(user_email, '@', 1));

    -- Crear perfil del usuario
    INSERT INTO public.profiles (id, name, email)
    VALUES (NEW.id, user_full_name, user_email);

    -- Crear cuentas de correo iniciales por defecto
    INSERT INTO public.user_accounts (user_id, email, is_primary)
    VALUES 
        (NEW.id, user_email, TRUE),
        (NEW.id, 'operaciones@triade.com', FALSE),
        (NEW.id, 'ventas@triade.com', FALSE)
    ON CONFLICT (user_id, email) DO NOTHING;

    -- Crear categorías iniciales por defecto
    INSERT INTO public.categories (user_id, name, color, keywords)
    VALUES
        (NEW.id, 'Trabajo', '#e63946', ARRAY['reunión', 'meeting', 'proyecto', 'project', 'informe', 'report']),
        (NEW.id, 'Finanzas', '#4c8bf5', ARRAY['factura', 'invoice', 'pago', 'payment', 'presupuesto', 'budget']),
        (NEW.id, 'Promociones', '#a06eff', ARRAY['oferta', 'offer', 'descuento', 'discount', 'promo']),
        (NEW.id, 'Social', '#3ccf91', ARRAY['invitación', 'invitation', 'evento', 'event', 'conexión']);

    RETURN NEW;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- Disparador en la tabla auth.users
DROP TRIGGER IF EXISTS on_auth_user_created ON auth.users;
CREATE TRIGGER on_auth_user_created
    AFTER INSERT ON auth.users
    FOR EACH ROW
    EXECUTE FUNCTION public.handle_new_user();

-- ---------------------------------------------------------------------
-- 9. ROW LEVEL SECURITY (RLS)
-- ---------------------------------------------------------------------
ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.user_accounts ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.categories ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.mails ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.attachments ENABLE ROW LEVEL SECURITY;

-- Políticas para profiles
CREATE POLICY "Usuarios pueden ver su propio perfil"
    ON public.profiles FOR SELECT
    USING (auth.uid() = id);

CREATE POLICY "Usuarios pueden actualizar su propio perfil"
    ON public.profiles FOR UPDATE
    USING (auth.uid() = id);

-- Políticas para user_accounts
CREATE POLICY "Usuarios pueden ver sus cuentas de correo"
    ON public.user_accounts FOR SELECT
    USING (auth.uid() = user_id);

CREATE POLICY "Usuarios pueden insertar sus cuentas de correo"
    ON public.user_accounts FOR INSERT
    WITH CHECK (auth.uid() = user_id);

CREATE POLICY "Usuarios pueden actualizar sus cuentas de correo"
    ON public.user_accounts FOR UPDATE
    USING (auth.uid() = user_id);

CREATE POLICY "Usuarios pueden eliminar sus cuentas de correo"
    ON public.user_accounts FOR DELETE
    USING (auth.uid() = user_id);

-- Políticas para categories
CREATE POLICY "Usuarios pueden ver sus categorías"
    ON public.categories FOR SELECT
    USING (auth.uid() = user_id);

CREATE POLICY "Usuarios pueden crear sus categorías"
    ON public.categories FOR INSERT
    WITH CHECK (auth.uid() = user_id);

CREATE POLICY "Usuarios pueden actualizar sus categorías"
    ON public.categories FOR UPDATE
    USING (auth.uid() = user_id);

CREATE POLICY "Usuarios pueden eliminar sus categorías"
    ON public.categories FOR DELETE
    USING (auth.uid() = user_id);

-- Políticas para mails
CREATE POLICY "Usuarios pueden ver sus correos"
    ON public.mails FOR SELECT
    USING (auth.uid() = user_id);

CREATE POLICY "Usuarios pueden crear correos"
    ON public.mails FOR INSERT
    WITH CHECK (auth.uid() = user_id);

CREATE POLICY "Usuarios pueden actualizar sus correos"
    ON public.mails FOR UPDATE
    USING (auth.uid() = user_id);

CREATE POLICY "Usuarios pueden eliminar sus correos"
    ON public.mails FOR DELETE
    USING (auth.uid() = user_id);

-- Políticas para attachments
CREATE POLICY "Usuarios pueden ver adjuntos de sus correos"
    ON public.attachments FOR SELECT
    USING (
        EXISTS (
            SELECT 1 FROM public.mails
            WHERE mails.id = attachments.mail_id
              AND mails.user_id = auth.uid()
        )
    );

CREATE POLICY "Usuarios pueden subir adjuntos a sus correos"
    ON public.attachments FOR INSERT
    WITH CHECK (
        EXISTS (
            SELECT 1 FROM public.mails
            WHERE mails.id = attachments.mail_id
              AND mails.user_id = auth.uid()
        )
    );

CREATE POLICY "Usuarios pueden eliminar adjuntos de sus correos"
    ON public.attachments FOR DELETE
    USING (
        EXISTS (
            SELECT 1 FROM public.mails
            WHERE mails.id = attachments.mail_id
              AND mails.user_id = auth.uid()
        )
    );

-- ---------------------------------------------------------------------
-- 10. ALMACENAMIENTO (SUPABASE STORAGE BUCKET)
-- ---------------------------------------------------------------------
INSERT INTO storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
VALUES (
    'mail-attachments',
    'mail-attachments',
    FALSE,
    52428800, -- 50MB
    ARRAY['image/*', 'application/pdf', 'application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'text/plain', 'application/zip']
)
ON CONFLICT (id) DO NOTHING;

CREATE POLICY "Acceso a adjuntos propios en storage"
    ON storage.objects FOR ALL
    USING (
        bucket_id = 'mail-attachments' AND
        auth.role() = 'authenticated' AND
        (storage.foldername(name))[1] = auth.uid()::text
    )
    WITH CHECK (
        bucket_id = 'mail-attachments' AND
        auth.role() = 'authenticated' AND
        (storage.foldername(name))[1] = auth.uid()::text
    );
