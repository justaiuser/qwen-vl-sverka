# -*- coding: utf-8 -*-
"""
Data models for accounting reconciliation.
"""
from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional
from datetime import date, datetime


class MatchStatus(str, Enum):
    EXACT = "EXACT"                         # Точное совпадение: сумма, номер и дата
    AMOUNT_AND_DOC = "AMOUNT_AND_DOC"       # Сумма и номер совпали (дата сдвинута)
    AMOUNT_AND_DATE = "AMOUNT_AND_DATE"     # Сумма и дата совпали (номер не указан/отличается)
    AMOUNT_DIFF = "AMOUNT_DIFF"             # Номер/дата сходятся, но сумма отличается
    ONLY_1C = "ONLY_1C"                     # Только в 1С (нет у контрагента)
    ONLY_COUNTERPARTY = "ONLY_CP"           # Только у контрагента (нет в 1С)


@dataclass
class ReconciliationItem:
    id: str
    source: str                             # '1c' или 'counterparty'
    date: Optional[date] = None
    date_str: str = ""
    doc_name: str = ""                      # Название документа (УПД, Акт, Платежное поручение)
    doc_number: str = ""                    # Номер документа (очищенный)
    doc_raw: str = ""                       # Исходная строка описания документа
    debit: float = 0.0                      # Дебет (руб.)
    credit: float = 0.0                     # Кредит (руб.)
    page: Optional[int] = None              # Номер страницы (для PDF/скана)
    line_no: Optional[int] = None           # Номер строки в таблице
    matched: bool = False                   # Флаг сопоставления

    @property
    def amount(self) -> float:
        """Основная сумма операции."""
        return self.debit if self.debit > 0 else self.credit

    @property
    def is_debit(self) -> bool:
        """True если операция по дебету."""
        return self.debit > 0


@dataclass
class ReconciliationStatement:
    source_name: str                        # '1C:Предприятие' или 'Контрагент'
    organization: str = ""                  # Наша организация
    counterparty: str = ""                  # Контрагент
    period_start: Optional[date] = None
    period_end: Optional[date] = None
    period_str: str = ""
    initial_balance: float = 0.0            # Положительное - Дебет (нам должны), Отрицательное - Кредит (мы должны)
    initial_balance_text: str = ""
    items: List[ReconciliationItem] = field(default_factory=list)
    turnover_debit: float = 0.0             # Обороты по дебету
    turnover_credit: float = 0.0            # Обороты по кредиту
    final_balance: float = 0.0              # Конечное сальдо
    final_balance_text: str = ""
    currency: str = "руб."


@dataclass
class MatchPair:
    id: str
    status: MatchStatus
    item_1c: Optional[ReconciliationItem] = None
    item_counterparty: Optional[ReconciliationItem] = None
    amount_1c: float = 0.0
    amount_counterparty: float = 0.0
    diff_amount: float = 0.0
    date_diff_days: Optional[int] = None
    explanation: str = ""

    @property
    def is_matched(self) -> bool:
        return self.status in (
            MatchStatus.EXACT,
            MatchStatus.AMOUNT_AND_DOC,
            MatchStatus.AMOUNT_AND_DATE
        )


@dataclass
class ReconciliationReport:
    statement_1c: ReconciliationStatement
    statement_counterparty: ReconciliationStatement
    pairs: List[MatchPair] = field(default_factory=list)
    
    # Итоговые метрики
    total_items_1c: int = 0
    total_items_cp: int = 0
    matched_count: int = 0
    discrepancy_count: int = 0
    
    # Расхождения сальдо и оборотов
    diff_initial_balance: float = 0.0
    diff_turnover_debit: float = 0.0
    diff_turnover_credit: float = 0.0
    diff_final_balance: float = 0.0
    
    has_discrepancies: bool = False
    summary_verdict: str = ""
