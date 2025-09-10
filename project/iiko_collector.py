# iiko_collector.py
import requests
import hashlib
import json
import urllib3
from pathlib import Path
from datetime import datetime

# Отключение предупреждений о сертификатах
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class IikoDataCollector:
    def __init__(self, base_url, login="angelinalina", password="092002"):
        self.base_url = base_url.strip()
        self.login = login
        self.password = password
        self.token = None
        self.session = requests.Session()
        self.session.verify = False

    def auth(self):
        """Аутентификация в системе iiko"""
        auth_url = f"{self.base_url}/auth"
        try:
            password_hash = hashlib.sha1(self.password.encode()).hexdigest()
            response = self.session.post(
                auth_url,
                data={'login': self.login, 'pass': password_hash},
                headers={'Content-Type': 'application/x-www-form-urlencoded'}
            )
            if response.status_code == 200:
                self.token = response.text.strip()
                return True
            return False
        except Exception as e:
            # print(f"Auth error: {e}") # Можно добавить логирование
            return False

    def get_report_data(self, preset_id, date_from, date_to, save_raw=False):
        """Получение данных отчета"""
        if not self.token and not self.auth():
            return None
        url = f"{self.base_url}/v2/reports/olap/byPresetId/{preset_id}"
        params = {
            'key': self.token,
            'dateFrom': date_from.strftime('%Y-%m-%d'),
            'dateTo': date_to.strftime('%Y-%m-%d')
        }
        try:
            response = self.session.get(url, params=params)
            if response.status_code == 200:
                json_data = response.json()
                if save_raw:
                    Path("raw_data").mkdir(exist_ok=True)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"raw_data/report_{timestamp}.json"
                    with open(filename, 'w', encoding='utf-8') as f:
                        json.dump(json_data, f, ensure_ascii=False, indent=2)
                return json_data
            return None
        except Exception as e:
            # print(f"Get report data error: {e}") # Можно добавить логирование
            return None

    def get_average_report_data(self, average_id, date_from, date_to):
        """Получение данных отчета по average_id"""
        if not self.token and not self.auth():
            return None
        url = f"{self.base_url}/v2/reports/olap/byPresetId/{average_id}"
        params = {
            'key': self.token,
            'dateFrom': date_from.strftime('%Y-%m-%d'),
            'dateTo': date_to.strftime('%Y-%m-%d')
        }
        try:
            response = self.session.get(url, params=params)
            if response.status_code == 200:
                return response.json()
            else:
                return None
        except Exception as e:
            # print(f"Get average report data error: {e}") # Можно добавить логирование
            return None