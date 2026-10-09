import os
from typing import List, Dict, Any, Optional
from dotenv import load_dotenv

load_dotenv(override=True)

# Demo / Sinov ma'lumotlari (Faqat Ism va Telefon raqami)
DEMO_PEOPLE = [
    {
        "id": "1",
        "full_name": "Karimov Sardor Baxtiyorovich",
        "phone_number": "+998901234567",
        "created_at": "2026-09-26 12:00:00"
    },
    {
        "id": "2",
        "full_name": "Toshmatov Jamshid Olimovich",
        "phone_number": "+998939876543",
        "created_at": "2026-09-26 12:05:00"
    },
    {
        "id": "3",
        "full_name": "Aliyeva Nigora Rustam qizi",
        "phone_number": "+998971112233",
        "created_at": "2026-09-26 12:10:00"
    },
    {
        "id": "4",
        "full_name": "Rahimov Aziz Akramovich",
        "phone_number": "+998998887766",
        "created_at": "2026-09-26 12:15:00"
    },
    {
        "id": "5",
        "full_name": "Usmonova Dilnoza Mansurovna",
        "phone_number": "+998945554433",
        "created_at": "2026-09-26 12:20:00"
    }
]

class SupabaseService:
    def __init__(self, url: Optional[str] = None, key: Optional[str] = None):
        self.url = url or os.getenv("SUPABASE_URL", "").strip()
        self.key = key or os.getenv("SUPABASE_KEY", "").strip()
        self.client = None
        self.is_connected = False
        self.table_exists = False
        self.status_message = ""
        self.init_client()

    def init_client(self):
        """Supabase mijozini ishga tushirish va ulanishni tekshirish"""
        load_dotenv(override=True)
        self.url = self.url or os.getenv("SUPABASE_URL", "").strip()
        self.key = self.key or os.getenv("SUPABASE_KEY", "").strip()
        if self.url and self.key and not self.url.startswith("https://your-project"):
            try:
                from supabase import create_client, Client
                self.client: Client = create_client(self.url, self.key)
                self.is_connected = True
                
                try:
                    res = self.client.table("people").select("id").limit(1).execute()
                    self.table_exists = True
                    self.status_message = "🟢 Supabase faol!"
                except Exception as table_err:
                    self.table_exists = False
                    self.status_message = f"⚠️ Supabase ogohlantirish: {table_err}"
            except Exception as e:
                self.client = None
                self.is_connected = False
                self.table_exists = False
                self.status_message = f"❌ Ulanishda xatolik: {e}"
        else:
            self.client = None
            self.is_connected = False
            self.table_exists = False
            self.status_message = "🟡 Sinov rejimi"

    def get_people(self, search_query: str = "") -> List[Dict[str, Any]]:
        """Odamlar ro'yxatini qidiruv bilan olish (Faqat FullName va PhoneNumber)"""
        if self.is_connected and self.table_exists and self.client:
            try:
                query = self.client.table("people").select("id, full_name, phone_number, created_at").order("created_at", desc=True)
                response = query.execute()
                people = response.data or []

                if search_query:
                    sq = search_query.lower()
                    people = [
                        p for p in people
                        if sq in str(p.get("full_name", "")).lower()
                        or sq in str(p.get("phone_number", "")).lower()
                    ]
                return people
            except Exception as e:
                print(f"Ma'lumotlarni olishda xatolik: {e}")
                return []
        else:
            filtered = DEMO_PEOPLE
            if search_query:
                sq = search_query.lower()
                filtered = [
                    p for p in filtered
                    if sq in str(p.get("full_name", "")).lower()
                    or sq in str(p.get("phone_number", "")).lower()
                ]
            return filtered

    def add_person(self, person_data: Dict[str, Any]) -> tuple[bool, str]:
        """Yangi odam qo'shish (Faqat FullName va PhoneNumber)"""
        clean_data = {
            "full_name": person_data.get("full_name", "").strip(),
            "phone_number": person_data.get("phone_number", "").strip()
        }
        if self.is_connected and self.client and self.table_exists:
            try:
                response = self.client.table("people").insert(clean_data).execute()
                if response.data:
                    return True, "Odam muvaffaqiyatli qo'shildi!"
                return False, "Kutilmagan xatolik yuz berdi."
            except Exception as e:
                err_str = str(e)
                if "violates not-null constraint" in err_str:
                    return False, "Supabase bazasida eski ustunlar qolgan. Iltimos, SQL Editor-da `ALTER TABLE public.people DROP COLUMN ...` skriptini ishga tushiring!"
                return False, f"Xatolik: {err_str}"
        else:
            import uuid
            from datetime import datetime
            new_item = {
                "id": str(uuid.uuid4()),
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                **clean_data
            }
            DEMO_PEOPLE.insert(0, new_item)
            return True, "Odam saqlandi (Sinov rejimi)."

    def update_person(self, person_id: str, updated_data: Dict[str, Any]) -> tuple[bool, str]:
        """Odam ma'lumotlarini yangilash"""
        clean_data = {
            "full_name": updated_data.get("full_name", "").strip(),
            "phone_number": updated_data.get("phone_number", "").strip()
        }
        if self.is_connected and self.client and self.table_exists:
            try:
                response = self.client.table("people").update(clean_data).eq("id", person_id).execute()
                if response.data:
                    return True, "Ma'lumotlar muvaffaqiyatli yangilandi!"
                return False, "Yangilash amalga oshmadi."
            except Exception as e:
                return False, f"Xatolik: {str(e)}"
        else:
            for p in DEMO_PEOPLE:
                if str(p.get("id")) == str(person_id):
                    p.update(clean_data)
                    return True, "Ma'lumotlar yangilandi (Sinov rejimi)."
            return False, "Odam topilmadi."

    def delete_person(self, person_id: str) -> tuple[bool, str]:
        """Odamni o'chirish"""
        if self.is_connected and self.client and self.table_exists:
            try:
                response = self.client.table("people").delete().eq("id", person_id).execute()
                return True, "Odam o'chirildi!"
            except Exception as e:
                return False, f"Xatolik: {str(e)}"
        else:
            global DEMO_PEOPLE
            initial_len = len(DEMO_PEOPLE)
            DEMO_PEOPLE = [p for p in DEMO_PEOPLE if str(p.get("id")) != str(person_id)]
            if len(DEMO_PEOPLE) < initial_len:
                return True, "Odam o'chirildi (Sinov rejimi)."
            return False, "Odam topilmadi."

    def delete_people_batch(self, person_ids: List[str]) -> tuple[int, str]:
        """Belgilangan odamlarni ommaviy o'chirish"""
        if not person_ids:
            return 0, "O'chirish uchun odamlar belgilanmagan."

        str_ids = [str(pid) for pid in person_ids if pid]
        if self.is_connected and self.client and self.table_exists:
            try:
                batch_size = 500
                total_deleted = 0
                for i in range(0, len(str_ids), batch_size):
                    chunk = str_ids[i:i + batch_size]
                    response = self.client.table("people").delete().in_("id", chunk).execute()
                    deleted_in_chunk = len(response.data) if response.data is not None else len(chunk)
                    total_deleted += deleted_in_chunk
                return total_deleted, f"{total_deleted} ta odam muvaffaqiyatli o'chirildi!"
            except Exception as e:
                return 0, f"O'chirishda xatolik: {str(e)}"
        else:
            global DEMO_PEOPLE
            initial_len = len(DEMO_PEOPLE)
            id_set = set(str_ids)
            DEMO_PEOPLE = [p for p in DEMO_PEOPLE if str(p.get("id")) not in id_set]
            deleted_count = initial_len - len(DEMO_PEOPLE)
            return deleted_count, f"{deleted_count} ta odam o'chirildi (Sinov rejimi)."

    def add_people_batch(self, people_list: List[Dict[str, Any]]) -> tuple[int, str]:
        """Excel orqali ommaviy yuklash"""
        if not people_list:
            return 0, "Yuklash uchun ma'lumotlar topilmadi."
        
        clean_list = [
            {
                "full_name": p.get("full_name", "").strip(),
                "phone_number": p.get("phone_number", "").strip()
            }
            for p in people_list if p.get("full_name") or p.get("phone_number")
        ]

        if self.is_connected and self.client and self.table_exists:
            try:
                # 1000 tadan bo'lib o'ta tez yuklaymiz (Supabase limiti va xatoliklarni oldini olish uchun)
                batch_size = 1000
                total_inserted = 0
                for i in range(0, len(clean_list), batch_size):
                    chunk = clean_list[i:i + batch_size]
                    response = self.client.table("people").insert(chunk).execute()
                    total_inserted += len(response.data) if response.data else len(chunk)
                return total_inserted, f"{total_inserted} ta odam muvaffaqiyatli saqlandi!"
            except Exception as e:
                err_str = str(e)
                if "violates not-null constraint" in err_str:
                    return 0, "Supabase bazasida eski ustunlar mavjud. Supabase SQL Editor-da `ALTER TABLE public.people DROP COLUMN ...` skriptini Run qiling!"
                return 0, f"Xatolik: {err_str}"
        else:
            import uuid
            from datetime import datetime
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for item in clean_list:
                new_item = {
                    "id": str(uuid.uuid4()),
                    "created_at": now_str,
                    **item
                }
                DEMO_PEOPLE.insert(0, new_item)
            return len(clean_list), f"{len(clean_list)} ta odam saqlandi (Sinov rejimi)."
