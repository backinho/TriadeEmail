-- =====================================================================
-- SEED DE PRUEBAS PARA SUPABASE (LOCAL DEV)
-- =====================================================================

-- Nota: Este script se ejecuta opcionalmente en entornos locales con `supabase db reset`
-- para pre-poblar datos de demostración en Triade Mail.

-- Ejemplo de datos de demostración (se asume que existe un usuario auth en local):
-- DO $$
-- DECLARE
--     demo_user_id UUID := '00000000-0000-0000-0000-000000000000';
-- BEGIN
--     INSERT INTO public.mails (user_id, from_name, from_email, to_address, subject, body, account, folder, unread, starred)
--     VALUES 
--         (demo_user_id, 'Halliburton', 'facturacion@halliburton.com', 'operaciones@triade.com', 'Factura Nº 4192 — VFD 200HP', 'Adjuntamos la factura del servicio de instalación…', 'operaciones@triade.com', 'inbox', true, false),
--         (demo_user_id, 'PDVSA Occidente', 'contacto@pdvsa.com', 'operaciones@triade.com', 'Reunión de coordinación proyecto Lago', 'Confirmamos reunión para revisar cronograma de rehabilitación…', 'operaciones@triade.com', 'inbox', true, true),
--         (demo_user_id, 'Schlumberger', 'ventas@slb.com', 'ventas@triade.com', 'Presupuesto BCP — 4 unidades', 'Enviamos el presupuesto solicitado para los equipos BCP…', 'ventas@triade.com', 'inbox', false, false);
-- END $$;
