# -*- coding: utf-8 -*-
"""
Client for interacting with local Qwen-VL via Ollama API.
"""
import io
import re
import json
import base64
import urllib.request
import urllib.error
from typing import Dict, Any, List, Optional
from PIL import Image


class QwenVLClient:
    def __init__(self, base_url: str = "http://127.0.0.1:11434", default_model: str = "qwen3-vl:4b-instruct"):
        self.base_url = base_url.rstrip("/")
        self.default_model = default_model

    def is_available(self) -> bool:
        """Проверка доступности Ollama сервера."""
        try:
            req = urllib.request.Request(f"{self.base_url}/api/tags")
            with urllib.request.urlopen(req, timeout=3) as resp:
                return resp.status == 200
        except Exception:
            return False

    def list_models(self) -> List[str]:
        """Возвращает список доступных моделей в Ollama."""
        try:
            req = urllib.request.Request(f"{self.base_url}/api/tags")
            with urllib.request.urlopen(req, timeout=5) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                return [m["name"] for m in data.get("models", [])]
        except Exception:
            return []

    def pil_to_base64(self, image: Image.Image, max_dimension: int = 1800) -> str:
        """Конвертирует изображение в оптимизированный base64 JPEG/PNG."""
        img = image.convert("RGB")
        w, h = img.size
        if max(w, h) > max_dimension:
            scale = max_dimension / max(w, h)
            new_size = (int(w * scale), int(h * scale))
            img = img.resize(new_size, Image.Resampling.LANCZOS)
        
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=90)
        return base64.b64encode(buf.getvalue()).decode("utf-8")

    def extract_act_from_image(
        self, 
        image: Image.Image, 
        model: Optional[str] = None,
        page_num: int = 1
    ) -> Dict[str, Any]:
        """
        Распознает акт сверки или страницу акта сверки с помощью Qwen-VL.
        Возвращает структурированный словарь с данными.
        """
        model_name = model or self.default_model
        b64_image = self.pil_to_base64(image)

        prompt = f"""Ты — профессиональный эксперт-бухгалтер по распознаванию актов сверки взаиморасчетов в РФ.
Перед тобой изображение страницы {page_num} акта сверки взаимных расчетов.

ВНИМАТЕЛЬНО изучи изображение и извлеки все табличные строки и реквизиты.
Обрати внимание на колонки:
- Дата операции (формат ДД.ММ.ГГГГ)
- Документ (название и номер: УПД, Товарная накладная, Акт, Платежное поручение, Счет-фактура)
- Дебет (сумма)
- Кредит (сумма)

Если в таблице две половины (По данным Организации 1 / По данным Организации 2), извлеки данные стороны контрагента, либо первой заполненной стороны.

Ответь СТРОГО в формате JSON без дополнительного текста:
{{
  "organization": "Название нашей организации или пусто",
  "counterparty": "Название контрагента или пусто",
  "period": "Период за который составлен акт, например 01.01.2024 - 31.12.2024",
  "initial_balance": {{
    "amount": 0.0,
    "side": "дебет или кредит или нет"
  }},
  "turnover_debit": 0.0,
  "turnover_credit": 0.0,
  "final_balance": {{
    "amount": 0.0,
    "side": "дебет или кредит или нет"
  }},
  "rows": [
    {{
      "date": "15.01.2024",
      "doc_type": "Товарная накладная",
      "doc_number": "125",
      "doc_full": "Товарная накладная № 125 от 15.01.2024",
      "debit": 54200.0,
      "credit": 0.0
    }}
  ]
}}

Правила сумм:
- Суммы должны быть числами (float). Разделитель точка. Пробелы удали.
- Если колонка Дебет или Кредит пустая, пиши 0.0.
- Не пропускай ни одной строки из таблицы операций.
"""

        payload = {
            "model": model_name,
            "prompt": prompt,
            "images": [b64_image],
            "stream": False,
            "format": "json",
            "options": {
                "temperature": 0.1,  # Низкая температура для максимальной точности фактов
                "num_ctx": 16384     # Достаточный контекст для больших таблиц
            }
        }

        req = urllib.request.Request(
            f"{self.base_url}/api/generate",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"}
        )

        try:
            with urllib.request.urlopen(req, timeout=180) as resp:
                result = json.loads(resp.read().decode("utf-8"))
                raw_text = result.get("response", "").strip()
                return self._parse_json_safe(raw_text)
        except urllib.error.HTTPError as e:
            err_body = e.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"Ошибка Ollama API ({e.code}): {err_body}")
        except Exception as e:
            raise RuntimeError(f"Ошибка соединения с Ollama: {str(e)}")

    def _parse_json_safe(self, text: str) -> Dict[str, Any]:
        """Безопасный парсинг JSON с удалением markdown-оберток."""
        text = text.strip()
        # Удаляем ```json ... ```
        if text.startswith("```"):
            lines = text.splitlines()
            if lines[0].startswith("```"):
                lines = lines[1:]
            if lines and lines[-1].startswith("```"):
                lines = lines[:-1]
            text = "\n".join(lines).strip()

        # Ищем первый { и последний }
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end != -1 and end > start:
            text = text[start:end+1]

        try:
            return json.loads(text)
        except json.JSONDecodeError:
            # Попытка исправить неэкранированные спецсимволы
            cleaned = re.sub(r'[\x00-\x1f]', ' ', text)
            try:
                return json.loads(cleaned)
            except Exception:
                return {"rows": [], "error": "Не удалось распарсить JSON ответ модели", "raw": text}
