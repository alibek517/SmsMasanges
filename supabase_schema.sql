-- ========================================================
-- Supabase SQL: Faqat FullName va PhoneNumber saqlanadigan yangi sxema
-- Supabase SQL Editor bo'limida ishga tushiring (Run)
-- ========================================================

-- 1-usul: Mavjud jadvaldan keraksiz ustunlarni o'chirish (Ma'lumotlar saqlanadi):
ALTER TABLE public.people 
    DROP COLUMN IF EXISTS district_name,
    DROP COLUMN IF EXISTS locality_name,
    DROP COLUMN IF EXISTS pinfl,
    DROP COLUMN IF EXISTS infl,
    DROP COLUMN IF EXISTS age,
    DROP COLUMN IF EXISTS document;

-- Indekslarni yangilash
DROP INDEX IF EXISTS idx_people_pinfl;
DROP INDEX IF EXISTS idx_people_district;
DROP INDEX IF EXISTS idx_people_locality;
CREATE INDEX IF NOT EXISTS idx_people_phone ON public.people (phone_number);
CREATE INDEX IF NOT EXISTS idx_people_fullname ON public.people (full_name);

-- RLS siyosati (To'liq ruxsat)
ALTER TABLE public.people ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Public full access policy" ON public.people;
CREATE POLICY "Public full access policy" ON public.people
    FOR ALL
    TO anon, authenticated
    USING (true)
    WITH CHECK (true);
