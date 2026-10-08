# -*- coding: utf-8 -*-
"""
Robust parser for 1C:Enterprise reconciliation statements in Excel (.xlsx, .xls).
"""
import re
from datetime import datetime, date
from typing import List, Optional, Tuple, Dict, Any
import openpyxl

from .models import ReconciliationItem, ReconciliationStatement


class Excel1CParser:
    def __init__(self):
        pass

    def parse(self, filepath: str) -> ReconciliationStatement:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheet = wb.active

        stmt = ReconciliationStatement(source_name="1C:Предприятие")

        # 1. Поиск метаданных в шапке (Организации, период, начальное сальдо)
        self._extract_header_metadata(sheet, stmt)

        # 2. Поиск заголовков колонок таблицы
        header_row, col_map = self._find_table_headers(sheet)

        if header_row is None or not col_map:
            # Попытка резервного эвристического поиска таблицы
            header_row, col_map = self._heuristic_table_search(sheet)

        if header_row is None:
            raise ValueError("Не удалось обнаружить структуру таблицы акта сверки в Excel файле.")

        # 3. Извлечение строк операций
        items = []
        max_row = sheet.max_row
        row_idx = header_row + 1

        while row_idx <= max_row:
            row_vals = [sheet.cell(row=row_idx, column=c).value for c in range(1, sheet.max_column + 1)]
            row_str = " ".join(str(v) for v in row_vals if v is not None).lower()

            # Проверка на начальное сальдо (должно проверяться ДО общего "сальдо на")
            if "сальдо начальное" in row_str or "нач. сальдо" in row_str or "входящее сальдо" in row_str:
                self._extract_initial_balance_row(sheet, row_idx, col_map, stmt, row_str)
                row_idx += 1
                continue

            # Проверка на окончание таблицы (Обороты, Сальдо конечное, Подписи)
            if "обороты за период" in row_str or "итого обороты" in row_str or "всего оборотов" in row_str:
                self._extract_turnovers(sheet, row_idx, col_map, stmt)
                row_idx += 1
                continue

            if "сальдо конечное" in row_str or "исходящее сальдо" in row_str or "сальдо на" in row_str:
                self._extract_final_balance(sheet, row_idx, col_map, stmt, row_str)
                row_idx += 1
                continue

            # Проверяем, является ли строка операцией
            item = self._parse_row(sheet, row_idx, col_map, len(items) + 1)
            if item:
                items.append(item)

            row_idx += 1

        stmt.items = items

        # Если обороты не были найдены в подвале, вычисляем их по строкам
        if stmt.turnover_debit == 0.0 and stmt.turnover_credit == 0.0 and items:
            stmt.turnover_debit = round(sum(i.debit for i in items), 2)
            stmt.turnover_credit = round(sum(i.credit for i in items), 2)

        return stmt

    def _extract_header_metadata(self, sheet, stmt: ReconciliationStatement):
        """Извлечение названий организаций и периода из первых 15 строк."""
        for r in range(1, min(20, sheet.max_row + 1)):
            for c in range(1, min(10, sheet.max_column + 1)):
                val = sheet.cell(row=r, column=c).value
                if not val or not isinstance(val, str):
                    continue
                
                # Поиск периода
                # Пример: "за период: 01.01.2024 - 31.12.2024" или "на 31.12.2024"
                m_period = re.search(r'(\d{2}\.\d{2}\.\d{4})\s*[-–—подо]\s*(\d{2}\.\d{2}\.\d{4})', val)
                if m_period:
                    stmt.period_str = f"{m_period.group(1)} - {m_period.group(2)}"
                    stmt.period_start = self._parse_date_str(m_period.group(1))
                    stmt.period_end = self._parse_date_str(m_period.group(2))

                # Поиск организаций
                # Пример: "между ООО "Ромашка" и ПАО "Лютик""
                m_parties = re.search(r'между\s+([«"][^"»]+[»"]|ООО\s+\S+|АО\s+\S+|ИП\s+\S+)\s+и\s+([«"][^"»]+[»"]|ООО\s+\S+|АО\s+\S+|ИП\s+\S+)', val, re.IGNORECASE)
                if m_parties:
                    stmt.organization = m_parties.group(1).strip()
                    stmt.counterparty = m_parties.group(2).strip()

                # Поиск начального сальдо в шапке
                if "сальдо начальное" in val.lower() or "сальдо на" in val.lower():
                    # Убираем дату, чтобы не спутать 01.01.2024 с суммой
                    val_no_date = re.sub(r'\d{2}\.\d{2}\.\d{4}', '', val)
                    m_bal = re.search(r'([\d\s]+[.,]\d{2})', val_no_date)
                    if m_bal:
                        amt = self._clean_number(m_bal.group(1))
                        stmt.initial_balance = amt
                        stmt.initial_balance_text = val.strip()

    def _find_table_headers(self, sheet) -> Tuple[Optional[int], Dict[str, int]]:
        """Поиск строки заголовков таблицы: Дата, Документ, Дебет, Кредит."""
        col_map = {}
        for r in range(1, min(30, sheet.max_row + 1)):
            row_map = {}
            for c in range(1, sheet.max_column + 1):
                val = sheet.cell(row=r, column=c).value
                if not val or not isinstance(val, str):
                    continue
                v_lower = val.lower().strip()
                if "дата" in v_lower and "date" not in row_map:
                    row_map["date"] = c
                elif ("документ" in v_lower or "содержание" in v_lower or "операция" in v_lower) and "doc" not in row_map:
                    row_map["doc"] = c
                elif "дебет" in v_lower and "debit" not in row_map:
                    row_map["debit"] = c
                elif "кредит" in v_lower and "credit" not in row_map:
                    row_map["credit"] = c

            # Если нашли все 4 ключевые колонки нашей стороны (или хотя бы Дата, Документ и Дебет/Кредит)
            if ("date" in row_map and "doc" in row_map) and ("debit" in row_map or "credit" in row_map):
                return r, row_map

        return None, {}

    def _heuristic_table_search(self, sheet) -> Tuple[Optional[int], Dict[str, int]]:
        """Резервный поиск: ищет первую строку, где есть дата и числовые суммы."""
        for r in range(1, min(40, sheet.max_row + 1)):
            date_col = None
            num_cols = []
            str_col = None
            for c in range(1, sheet.max_column + 1):
                val = sheet.cell(row=r, column=c).value
                if isinstance(val, (datetime, date)):
                    date_col = c
                elif isinstance(val, str) and re.match(r'^\d{2}\.\d{2}\.\d{2,4}$', val.strip()):
                    date_col = c
                elif isinstance(val, (int, float)) and val > 0:
                    num_cols.append(c)
                elif isinstance(val, str) and len(val.strip()) > 3:
                    if not str_col:
                        str_col = c

            if date_col and (num_cols or str_col):
                # Найдена первая строка данных, значит шапка на строку выше
                header_r = max(1, r - 1)
                col_map = {
                    "date": date_col,
                    "doc": str_col if str_col else (date_col + 1),
                    "debit": num_cols[0] if len(num_cols) > 0 else date_col + 2,
                    "credit": num_cols[1] if len(num_cols) > 1 else date_col + 3,
                }
                return header_r, col_map

        return None, {}

    def _parse_row(self, sheet, row_idx: int, col_map: Dict[str, int], line_num: int) -> Optional[ReconciliationItem]:
        """Парсинг одной строки операции."""
        date_val = sheet.cell(row=row_idx, column=col_map["date"]).value
        doc_val = sheet.cell(row=row_idx, column=col_map.get("doc", 0)).value if "doc" in col_map else ""
        debit_val = sheet.cell(row=row_idx, column=col_map.get("debit", 0)).value if "debit" in col_map else 0
        credit_val = sheet.cell(row=row_idx, column=col_map.get("credit", 0)).value if "credit" in col_map else 0

        parsed_date, date_str = self._resolve_date(date_val)
        debit_amt = self._clean_number(debit_val)
        credit_amt = self._clean_number(credit_val)

        # Если нет ни даты, ни сумм - строка пустая или разделитель
        if not date_str and debit_amt == 0.0 and credit_amt == 0.0:
            return None

        # Если есть сумма, но нет даты, проверяем не заголовок ли это
        doc_str = str(doc_val).strip() if doc_val is not None else ""
        if not date_str and (debit_amt > 0 or credit_amt > 0):
            # Возможно, промежуточный итог
            if any(w in doc_str.lower() for w in ["итого", "сальдо", "обороты"]):
                return None

        # Извлечение номера и типа документа
        doc_type, doc_num = self._extract_doc_info(doc_str)

        return ReconciliationItem(
            id=f"1c_{line_num}_{row_idx}",
            source="1c",
            date=parsed_date,
            date_str=date_str,
            doc_name=doc_type,
            doc_number=doc_num,
            doc_raw=doc_str,
            debit=debit_amt,
            credit=credit_amt,
            line_no=line_num
        )

    def _resolve_date(self, val: Any) -> Tuple[Optional[date], str]:
        if val is None:
            return None, ""
        if isinstance(val, datetime):
            return val.date(), val.strftime("%d.%m.%Y")
        if isinstance(val, date):
            return val, val.strftime("%d.%m.%Y")
        
        s = str(val).strip()
        m = re.search(r'(\d{2})[./\-](\d{2})[./\-](\d{2,4})', s)
        if m:
            d, mth, y = m.group(1), m.group(2), m.group(3)
            if len(y) == 2:
                y = "20" + y
            try:
                dt = date(int(y), int(mth), int(d))
                return dt, dt.strftime("%d.%m.%Y")
            except Exception:
                return None, f"{d}.{mth}.{y}"
        return None, ""

    def _clean_number(self, val: Any) -> float:
        if val is None or val == "":
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        
        # Строка с числом
        s = str(val).replace('\xa0', '').replace(' ', '').replace(',', '.')
        m = re.search(r'[-+]?\d*\.?\d+', s)
        if m:
            try:
                return float(m.group(0))
            except ValueError:
                return 0.0
        return 0.0

    def _extract_doc_info(self, raw: str) -> Tuple[str, str]:
        """Извлечение типа и номера документа из строки 1С."""
        if not raw:
            return "", ""
        
        doc_type = ""
        doc_num = ""

        # Типичные типы документов 1С
        types = [
            "Реализация (акт, накладная, УПД)",
            "Реализация товаров и услуг",
            "Поступление (акт, накладная, УПД)",
            "Поступление товаров и услуг",
            "Поступление на расчетный счет",
            "Списание с расчетного счета",
            "Платежное поручение",
            "Товарная накладная",
            "Счет-фактура",
            "Акт сверки",
            "УПД",
            "Акт"
        ]
        for t in types:
            if t.lower() in raw.lower():
                doc_type = t
                break

        # Поиск номера документа: № 123, №123/45, № ПР-000012
        m_num = re.search(r'[№#N]\s*([A-Za-zА-Яа-я0-9_\-\/]+)', raw)
        if m_num:
            doc_num = m_num.group(1).lstrip("0") or "0"
        else:
            # Попытка найти отдельно стоящий номер перед 'от'
            m_alt = re.search(r'(?:документ|накладная|акт|упд|пп|п/п)\s+([A-Za-zА-Яа-я0-9_\-\/]+)\s+от', raw, re.IGNORECASE)
            if m_alt:
                doc_num = m_alt.group(1).lstrip("0") or "0"

        if not doc_type:
            doc_type = raw.split("№")[0].strip() if "№" in raw else raw

        return doc_type, doc_num

    def _extract_turnovers(self, sheet, row_idx: int, col_map: Dict[str, int], stmt: ReconciliationStatement):
        """Извлечение оборотов."""
        if "debit" in col_map:
            stmt.turnover_debit = self._clean_number(sheet.cell(row=row_idx, column=col_map["debit"]).value)
        if "credit" in col_map:
            stmt.turnover_credit = self._clean_number(sheet.cell(row=row_idx, column=col_map["credit"]).value)

    def _extract_final_balance(self, sheet, row_idx: int, col_map: Dict[str, int], stmt: ReconciliationStatement, row_str: str):
        """Извлечение конечного сальдо."""
        stmt.final_balance_text = row_str
        # Проверяем колонки дебет/кредит
        d_val = self._clean_number(sheet.cell(row=row_idx, column=col_map.get("debit", 0)).value) if "debit" in col_map else 0
        c_val = self._clean_number(sheet.cell(row=row_idx, column=col_map.get("credit", 0)).value) if "credit" in col_map else 0
        if d_val > 0:
            stmt.final_balance = d_val
        elif c_val > 0:
            stmt.final_balance = -c_val
        else:
            m = re.search(r'([\d\s]+[.,]\d{2})', row_str)
            if m:
                stmt.final_balance = self._clean_number(m.group(1))

    def _extract_initial_balance_row(self, sheet, row_idx: int, col_map: Dict[str, int], stmt: ReconciliationStatement, row_str: str):
        stmt.initial_balance_text = row_str
        d_val = self._clean_number(sheet.cell(row=row_idx, column=col_map.get("debit", 0)).value) if "debit" in col_map else 0
        c_val = self._clean_number(sheet.cell(row=row_idx, column=col_map.get("credit", 0)).value) if "credit" in col_map else 0
        if d_val > 0:
            stmt.initial_balance = d_val
        elif c_val > 0:
            stmt.initial_balance = -c_val

    def _parse_date_str(self, s: str) -> Optional[date]:
        try:
            return datetime.strptime(s.strip(), "%d.%m.%Y").date()
        except Exception:
            return None
