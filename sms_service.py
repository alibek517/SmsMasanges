import os
import time
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from typing import List, Dict, Any, Tuple
from dotenv import load_dotenv

load_dotenv(override=True)

DEFAULT_SENDER = "+998937078047"

class SMSService:
    def __init__(self):
        load_dotenv(override=True)
        self.sender_number = os.getenv("SMS_SENDER_NUMBER", DEFAULT_SENDER).strip()
        self.gateway_type = os.getenv("SMS_GATEWAY_TYPE", "infinireach").strip() # 'infinireach', 'android', 'eskiz'
        self.gateway_url = os.getenv("SMS_GATEWAY_URL", "https://api.infinireach.io").strip()
        self.gateway_token = os.getenv("SMS_GATEWAY_TOKEN", "").strip()
        self.gateway_login = os.getenv("SMS_GATEWAY_LOGIN", "").strip()
        self.gateway_password = os.getenv("SMS_GATEWAY_PASSWORD", "").strip()
        self.eskiz_token = os.getenv("ESKIZ_TOKEN", "").strip()
        self.history: List[Dict[str, Any]] = []

        # Bardoshli HTTP sessiya (Connection pooling & avto-qayta urinish)
        self.session = requests.Session()
        retries = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504], allowed_methods=["POST", "GET"])
        adapter = HTTPAdapter(max_retries=retries, pool_connections=10, pool_maxsize=20)
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)

    def is_configured(self) -> bool:
        """Real SMS shlyuzi sozlanganmi?"""
        if self.gateway_token:
            return True
        if self.gateway_url and (self.gateway_login and self.gateway_password):
            return True
        if self.eskiz_token or (self.gateway_login and "@" in self.gateway_login):
            return True
        return False

    def save_config(self, gateway_type: str = None, url: str = None, token: str = None, login: str = None, password: str = None, sender_number: str = None):
        """Sozlamalarni .env fayliga toza saqlash"""
        env_path = ".env"
        data = {}
        if os.path.exists(env_path):
            with open(env_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if "=" in line and not line.startswith("#"):
                        k, v = line.split("=", 1)
                        data[k.strip()] = v.strip()

        if gateway_type is not None:
            self.gateway_type = gateway_type
            data["SMS_GATEWAY_TYPE"] = gateway_type
        if url is not None:
            self.gateway_url = url
            data["SMS_GATEWAY_URL"] = url
        if token is not None:
            self.gateway_token = token
            data["SMS_GATEWAY_TOKEN"] = token
        if login is not None:
            self.gateway_login = login
            data["SMS_GATEWAY_LOGIN"] = login
        if password is not None:
            self.gateway_password = password
            data["SMS_GATEWAY_PASSWORD"] = password
        if sender_number is not None:
            self.sender_number = sender_number.strip()
            data["SMS_SENDER_NUMBER"] = sender_number.strip()

        with open(env_path, "w", encoding="utf-8") as f:
            for k, v in data.items():
                f.write(f"{k}={v}\n")

    def format_message(self, template: str, person: Dict[str, Any]) -> str:
        """Shablon ichidagi {FullName} o'rniga haqiqiy ismni qo'yish"""
        full_name = person.get("full_name", "").strip()
        phone_number = person.get("phone_number", "").strip()
        
        msg = template
        for tag in ["{FullName}", "{fullname}", "{FULLNAME}", "{Ism}", "{ism}"]:
            msg = msg.replace(tag, full_name)
        for tag in ["{PhoneNumberi}", "{phonenumberi}", "{Telefon}", "{telefon}"]:
            msg = msg.replace(tag, phone_number)
            
        return msg

    def send_via_infinireach(self, phone: str, text: str, sender: str = None) -> Tuple[bool, str]:
        """InfiniReach Cloud API orqali ulangan telefondagi SIM kartadan SMS jo'natish"""
        api_key = self.gateway_token or self.gateway_password
        if not api_key:
            return False, "InfiniReach API Key kiritilmagan! Sozlamalardan API Key-ni kiriting."

        clean_phone = phone.strip().replace(" ", "").replace("-", "")
        if not clean_phone.startswith("+"):
            clean_phone = "+" + clean_phone

        sender_from = sender or self.sender_number
        if not sender_from.startswith("+"):
            sender_from = "+" + sender_from

        headers = {
            "Content-Type": "application/json",
            "X-API-Key": api_key,
            "Authorization": f"Bearer {api_key}"
        }

        # 1-urinish: InfiniReach v1 messages endpoint
        url_v1 = "https://api.infinireach.io/api/v1/messages"
        payload_v1 = {
            "to": clean_phone,
            "message": text,
            "from": sender_from,
            "channel": "sms"
        }

        try:
            res = self.session.post(url_v1, json=payload_v1, headers=headers, timeout=15)
            data = res.json() if res.content else {}

            if res.status_code in [200, 201, 202]:
                if data.get("success") is False or data.get("status") == "failed":
                    err_msg = data.get("errorMessage") or data.get("error") or data.get("errorCode") or "Xatolik yuz berdi"
                    if "SIM_SLOT_NOT_FOUND" in str(err_msg) or "not active" in str(err_msg):
                        return False, "Telefonda SIM-karta ulanishi uzilgan! Telefondagi SMS Gateway ilovasini oching yoki batareya cheklovini olib tashlang."
                    return False, f"InfiniReach xatosi: {err_msg}"
                return True, "SMS telefoningizdagi SIM kartadan muvaffaqiyatli jo'natildi!"

            err_text = data.get("errorMessage") or data.get("error") or res.text[:120]
            return False, f"InfiniReach xatosi ({res.status_code}): {err_text}"
        except Exception as e:
            return False, f"InfiniReach ulanish xatosi: {str(e)}"

    def send_via_android_gateway(self, phone: str, text: str) -> Tuple[bool, str]:
        """Lokal Android SMS Gateway orqali jo'natish"""
        clean_phone = phone.strip().replace(" ", "").replace("-", "")
        url = self.gateway_url.rstrip("/")
        
        endpoints = [
            f"{url}/send",
            f"{url}/message",
            f"{url}/3rdparty/v1/messages",
            url
        ]

        headers = {"Content-Type": "application/json"}
        auth = None

        if self.gateway_token:
            headers["Authorization"] = f"Bearer {self.gateway_token}"
            headers["X-API-Key"] = self.gateway_token
        elif self.gateway_login and self.gateway_password:
            auth = (self.gateway_login, self.gateway_password)

        payload_variants = [
            {"to": clean_phone, "message": text},
            {"textMessage": {"text": text}, "phoneNumbers": [clean_phone]}
        ]

        last_error = ""
        for ep in endpoints:
            for payload in payload_variants:
                try:
                    res = requests.post(ep, json=payload, headers=headers, auth=auth, timeout=8)
                    if res.status_code in [200, 201, 202]:
                        return True, "SMS jo'natildi!"
                    last_error = f"Status {res.status_code}: {res.text[:80]}"
                except Exception as e:
                    last_error = str(e)
                    break

        return False, f"SMS Gateway xatosi: {last_error}"

    def send_via_eskiz(self, phone: str, text: str) -> Tuple[bool, str]:
        """Eskiz.uz API orqali jo'natish"""
        clean_phone = phone.strip().replace("+", "").replace(" ", "").replace("-", "")
        
        # Token olish (agar yo'q bo'lsa)
        if not self.eskiz_token and self.gateway_login and self.gateway_password:
            try:
                auth_res = requests.post("https://notify.eskiz.uz/api/auth/login", data={
                    "email": self.gateway_login,
                    "password": self.gateway_password
                }, timeout=10)
                if auth_res.status_code == 200:
                    self.eskiz_token = auth_res.json().get("data", {}).get("token", "")
            except Exception as e:
                return False, f"Eskiz login xatosi: {e}"

        if not self.eskiz_token:
            return False, "Eskiz tokeni mavjud emas"

        try:
            headers = {"Authorization": f"Bearer {self.eskiz_token}"}
            res = requests.post("https://notify.eskiz.uz/api/message/sms/send", headers=headers, data={
                "mobile_phone": clean_phone,
                "message": text,
                "from": "4546"
            }, timeout=10)
            if res.status_code == 200:
                return True, "SMS Eskiz orqali yuborildi"
            return False, f"Eskiz xatosi: {res.text[:100]}"
        except Exception as e:
            return False, f"Eskiz ulanish xatosi: {e}"

    def send_single_sms(self, recipient_phone: str, message_text: str, sender: str = None) -> Tuple[bool, str]:
        """Bitta raqamga SMS jo'natish"""
        # 1. Agar Lokal Android Gateway tanlangan bo'lsa
        if self.gateway_type == "android" or (self.gateway_url and "192.168" in self.gateway_url):
            return self.send_via_android_gateway(recipient_phone, message_text)

        # 2. Agar InfiniReach yoki Token kiritilgan bo'lsa
        if self.gateway_token:
            if "infinireach" in self.gateway_url.lower() or not self.gateway_url:
                return self.send_via_infinireach(recipient_phone, message_text, sender=sender)
            else:
                return self.send_via_android_gateway(recipient_phone, message_text)

        # 3. Agar Gateway URL ko'rsatilgan bo'lsa
        if self.gateway_url and "eskiz" not in self.gateway_url.lower():
            if "infinireach" in self.gateway_url.lower():
                return self.send_via_infinireach(recipient_phone, message_text, sender=sender)
            return self.send_via_android_gateway(recipient_phone, message_text)
        
        # 4. Agar Eskiz sozlamalari kiritilgan bo'lsa
        if "eskiz" in self.gateway_url.lower() or (self.gateway_login and "@" in self.gateway_login) or self.eskiz_token:
            return self.send_via_eskiz(recipient_phone, message_text)

        # 5. Agar Gateway hali ulanmagan bo'lsa
        return False, "Telefon (Lokal SMS Gateway yoki API Key) ulanmagan!"

    def send_bulk_sms(
        self,
        recipients: List[Dict[str, Any]],
        message_template: str,
        sender: str = None
    ) -> List[Dict[str, Any]]:
        """Tanlangan barcha odamlarga ketma-ket yuborish"""
        results = []
        sender_num = sender or self.sender_number
        from datetime import datetime

        for person in recipients:
            personalized_text = self.format_message(message_template, person)
            phone = person.get("phone_number", "")
            name = person.get("full_name", "")

            success, status_msg = self.send_single_sms(phone, personalized_text, sender=sender_num)
            
            record = {
                "name": name,
                "phone": phone,
                "message": personalized_text,
                "sender": sender_num,
                "success": success,
                "status": status_msg,
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            results.append(record)
            self.history.insert(0, record)

        return results

    def get_message_logs(self, limit: int = 100) -> List[Dict[str, Any]]:
        """InfiniReach Cloud API orqali barcha SMS xabarlar tarixini (Message Logs) olish"""
        api_key = self.gateway_token or self.gateway_password
        if not api_key:
            return []

        headers = {
            "Content-Type": "application/json",
            "X-API-Key": api_key,
            "Authorization": f"Bearer {api_key}"
        }

        try:
            url = f"https://api.infinireach.io/api/v1/messages?limit={limit}"
            res = self.session.get(url, headers=headers, timeout=15)
            if res.status_code == 200:
                data = res.json()
                return data.get("messages", [])
        except Exception as e:
            print("Message logs fetch error:", e)

        return []

    def get_devices(self) -> List[Dict[str, Any]]:
        """InfiniReach Cloud API orqali ulangan qurilmalar ro'yxatini olish"""
        api_key = self.gateway_token or self.gateway_password
        if not api_key:
            return []

        headers = {
            "Content-Type": "application/json",
            "X-API-Key": api_key,
            "Authorization": f"Bearer {api_key}"
        }

        try:
            url = "https://api.infinireach.io/api/v1/devices"
            res = self.session.get(url, headers=headers, timeout=10)
            if res.status_code == 200:
                data = res.json()
                return data.get("devices", [])
        except Exception as e:
            print("Devices fetch error:", e)

        return []

