-- =====================================================================
-- MIGRACIÓN 20260727000000: TABLA Y RELACIÓN DE CATEGORÍAS ÚNICAS POR USUARIO
-- =====================================================================

-- 1. Asegurar estructura completa de la tabla categories con FK a profiles(id) y UNIQUE(user_id, name)
CREATE TABLE IF NOT EXISTS public.categories (
    id UUID NOT NULL DEFAULT gen_random_uuid(),
    user_id UUID NOT NULL,
    name TEXT NOT NULL,
    color TEXT NOT NULL DEFAULT '#e63946'::text,
    keywords TEXT[] NOT NULL DEFAULT '{}'::text[],
    created_at TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT timezone('utc'::text, now()),
    CONSTRAINT categories_pkey PRIMARY KEY (id),
    CONSTRAINT categories_user_id_name_key UNIQUE (user_id, name),
    CONSTRAINT categories_user_id_fkey FOREIGN KEY (user_id) REFERENCES public.profiles(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_categories_user_id ON public.categories USING btree (user_id);

-- 2. Actualizar función handle_new_user para NO insertar categorías ni cuentas por defecto
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
    VALUES (NEW.id, user_full_name, user_email)
    ON CONFLICT (id) DO NOTHING;

    RETURN NEW;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- 3. Habilitar políticas RLS para public.categories
ALTER TABLE public.categories ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "Usuarios pueden ver sus categorías" ON public.categories;
CREATE POLICY "Usuarios pueden ver sus categorías"
    ON public.categories FOR SELECT
    USING (auth.uid() = user_id);

DROP POLICY IF EXISTS "Usuarios pueden crear sus categorías" ON public.categories;
CREATE POLICY "Usuarios pueden crear sus categorías"
    ON public.categories FOR INSERT
    WITH CHECK (auth.uid() = user_id);

DROP POLICY IF EXISTS "Usuarios pueden actualizar sus categorías" ON public.categories;
CREATE POLICY "Usuarios pueden actualizar sus categorías"
    ON public.categories FOR UPDATE
    USING (auth.uid() = user_id);

DROP POLICY IF EXISTS "Usuarios pueden eliminar sus categorías" ON public.categories;
CREATE POLICY "Usuarios pueden eliminar sus categorías"
    ON public.categories FOR DELETE
    USING (auth.uid() = user_id);
