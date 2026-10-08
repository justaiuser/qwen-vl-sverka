# -*- coding: utf-8 -*-
"""
Hybrid parser for Counterparty acts: Supports electronic PDFs, scanned PDFs, and images using Qwen-VL.
"""
import re
from datetime import datetime, date
from typing import List, Optional, Tuple, Dict, Any
from pathlib import Path
from PIL import Image

import pymupdf  # PyMuPDF

from .models import ReconciliationItem, ReconciliationStatement
from .qwen_client import QwenVLClient


class PDFCounterpartyParser:
    def __init__(self, qwen_client: Optional[QwenVLClient] = None, model_name: str = "qwen3-vl:4b-instruct"):
        self.qwen_client = qwen_client or QwenVLClient()
        self.model_name = model_name

    def parse(self, filepath: str, force_ocr: bool = False) -> ReconciliationStatement:
        path = Path(filepath)
        ext = path.suffix.lower()

        stmt = ReconciliationStatement(source_name="Контрагент")

        if ext in (".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".webp"):
            # Одиночное изображение
            img = Image.open(filepath)
            self._parse_with_qwen(stmt, [img])
            return stmt

        if ext == ".pdf":
            doc = pymupdf.open(filepath)
            num_pages = len(doc)

            # Проверяем наличие текстового слоя
            total_text_len = sum(len(page.get_text().strip()) for page in doc)
            is_scanned = total_text_len < 100 or force_ocr

            if not is_scanned:
                # Попытка извлечь текст напрямую из цифрового PDF
                success = self._parse_digital_pdf(doc, stmt)
                if success and len(stmt.items) > 0:
                    return stmt

            # Если PDF сканированный или прямой парсинг не дал строк — используем Qwen-VL
            page_images = []
            for page in doc:
                pix = page.get_pixmap(dpi=200)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                page_images.append(img)

            self._parse_with_qwen(stmt, page_images)
            return stmt

        raise ValueError(f"Неподдерживаемый формат файла: {ext}")

    def _parse_digital_pdf(self, doc, stmt: ReconciliationStatement) -> bool:
        """Быстрый парсер цифрового PDF с текстовым слоем."""
        all_lines = []
        for page in doc:
            text = page.get_text("text")
            all_lines.extend(text.splitlines())

        # Ищем шапку и строки операций
        items = []
        for line in all_lines:
            line_str = line.strip()
            if not line_str:
                continue

            # Поиск периода
            m_p = re.search(r'(\d{2}\.\d{2}\.\d{4})\s*[-–—подо]\s*(\d{2}\.\d{2}\.\d{4})', line_str)
            if m_p and not stmt.period_str:
                stmt.period_str = f"{m_p.group(1)} - {m_p.group(2)}"

            # Проверка строки таблицы: должна содержать дату и сумму
            # Пример: "15.01.2024 Товарная накладная № 125 54 200.00 0.00"
            m_date = re.search(r'(\d{2}\.\d{2}\.\d{4})', line_str)
            if m_date:
                # Поиск числовых сумм в конце строки
                # Ищем числа с разделителями
                tokens = line_str.split()
                # Извлекаем все суммы
                amounts = []
                for token in reversed(tokens):
                    clean_t = token.replace('\xa0', '').replace(',', '.')
                    if re.match(r'^\d+(\.\d{2})?$', clean_t):
                        try:
                            amounts.append(float(clean_t))
                        except ValueError:
                            pass
                        if len(amounts) == 2:
                            break

                if amounts:
                    amounts.reverse()
                    debit = amounts[0] if len(amounts) >= 1 else 0.0
                    credit = amounts[1] if len(amounts) >= 2 else 0.0
                    
                    dt_obj = self._parse_date_str(m_date.group(1))
                    doc_desc = line_str[m_date.end():].strip()
                    # Убираем числа из описания документа
                    doc_desc = re.sub(r'[\d\s.,]+$', '', doc_desc).strip()
                    doc_type, doc_num = self._extract_doc_info(doc_desc)

                    items.append(ReconciliationItem(
                        id=f"cp_pdf_{len(items)+1}",
                        source="counterparty",
                        date=dt_obj,
                        date_str=m_date.group(1),
                        doc_name=doc_type,
                        doc_number=doc_num,
                        doc_raw=doc_desc,
                        debit=debit,
                        credit=credit,
                        line_no=len(items)+1
                    ))

        if items:
            stmt.items = items
            stmt.turnover_debit = round(sum(i.debit for i in items), 2)
            stmt.turnover_credit = round(sum(i.credit for i in items), 2)
            return True
        return False

    def _parse_with_qwen(self, stmt: ReconciliationStatement, images: List[Image.Image]):
        """Распознавание скана с помощью Qwen-VL через Ollama."""
        all_items = []

        for p_idx, img in enumerate(images, start=1):
            res_data = self.qwen_client.extract_act_from_image(
                image=img,
                model=self.model_name,
                page_num=p_idx
            )

            if p_idx == 1:
                # Извлекаем метаданные с первой страницы
                if res_data.get("organization"):
                    stmt.organization = res_data["organization"]
                if res_data.get("counterparty"):
                    stmt.counterparty = res_data["counterparty"]
                if res_data.get("period"):
                    stmt.period_str = res_data["period"]
                
                init_b = res_data.get("initial_balance", {})
                if isinstance(init_b, dict):
                    amt = float(init_b.get("amount", 0.0) or 0.0)
                    side = str(init_b.get("side", "")).lower()
                    stmt.initial_balance = amt if "деб" in side else (-amt if "кред" in side else amt)

            # Строки операций со страницы
            rows = res_data.get("rows", [])
            for r in rows:
                if not isinstance(r, dict):
                    continue
                
                date_str = str(r.get("date", "")).strip()
                dt_obj = self._parse_date_str(date_str)

                debit = float(r.get("debit", 0.0) or 0.0)
                credit = float(r.get("credit", 0.0) or 0.0)
                
                # Пропускаем строки без сумм
                if debit == 0.0 and credit == 0.0:
                    continue

                doc_full = str(r.get("doc_full", "") or r.get("document", "")).strip()
                doc_type = str(r.get("doc_type", "")).strip()
                doc_num = str(r.get("doc_number", "")).strip()

                if not doc_num and doc_full:
                    m_num = re.search(r'[№#N]\s*([A-Za-zА-Яа-я0-9_\-\/]+)', doc_full)
                    if m_num:
                        doc_num = m_num.group(1).lstrip("0") or "0"

                if not doc_type and doc_full:
                    doc_type, _ = self._extract_doc_info(doc_full)

                item = ReconciliationItem(
                    id=f"cp_scan_p{p_idx}_{len(all_items)+1}",
                    source="counterparty",
                    date=dt_obj,
                    date_str=date_str,
                    doc_name=doc_type,
                    doc_number=doc_num,
                    doc_raw=doc_full,
                    debit=debit,
                    credit=credit,
                    page=p_idx,
                    line_no=len(all_items) + 1
                )
                all_items.append(item)

            # Если на последней странице есть обороты или конечное сальдо
            if p_idx == len(images):
                if res_data.get("turnover_debit"):
                    stmt.turnover_debit = float(res_data["turnover_debit"])
                if res_data.get("turnover_credit"):
                    stmt.turnover_credit = float(res_data["turnover_credit"])
                
                fin_b = res_data.get("final_balance", {})
                if isinstance(fin_b, dict):
                    amt = float(fin_b.get("amount", 0.0) or 0.0)
                    side = str(fin_b.get("side", "")).lower()
                    stmt.final_balance = amt if "деб" in side else (-amt if "кред" in side else amt)

        stmt.items = all_items

        if stmt.turnover_debit == 0.0 and stmt.turnover_credit == 0.0 and all_items:
            stmt.turnover_debit = round(sum(i.debit for i in all_items), 2)
            stmt.turnover_credit = round(sum(i.credit for i in all_items), 2)

    def _parse_date_str(self, s: str) -> Optional[date]:
        if not s:
            return None
        m = re.search(r'(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})', s)
        if m:
            d, mth, y = m.group(1), m.group(2), m.group(3)
            if len(y) == 2:
                y = "20" + y
            try:
                return date(int(y), int(mth), int(d))
            except Exception:
                return None
        return None

    def _extract_doc_info(self, raw: str) -> Tuple[str, str]:
        if not raw:
            return "", ""
        doc_type = ""
        doc_num = ""
        types = [
            "Товарная накладная", "УПД", "Платежное поручение", "Акт", 
            "Счет-фактура", "Поступление", "Реализация"
        ]
        for t in types:
            if t.lower() in raw.lower():
                doc_type = t
                break
        m_num = re.search(r'[№#N]\s*([A-Za-zА-Яа-я0-9_\-\/]+)', raw)
        if m_num:
            doc_num = m_num.group(1).lstrip("0") or "0"
        return doc_type, doc_num
