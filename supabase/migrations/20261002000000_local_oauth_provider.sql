ALTER TABLE public.user_accounts
    ADD COLUMN IF NOT EXISTS provider TEXT CHECK (provider IN ('gmail', 'outlook', 'custom'));

UPDATE public.user_accounts
SET provider = CASE
    WHEN lower(email) LIKE '%@gmail.com' OR lower(email) LIKE '%@googlemail.com' THEN 'gmail'
    WHEN lower(email) LIKE '%@outlook.%' OR lower(email) LIKE '%@hotmail.%' OR lower(email) LIKE '%@live.%' THEN 'outlook'
    ELSE provider
END
WHERE provider IS NULL;

UPDATE public.user_accounts
SET access_token = NULL,
    refresh_token = NULL,
    expires_at = NULL
WHERE lower(email) LIKE '%@gmail.com'
   OR lower(email) LIKE '%@googlemail.com';