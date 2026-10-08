# -*- coding: utf-8 -*-
"""
Smart reconciliation matching engine.
Applies Russian accounting rules:
- 1C Debit <-> Counterparty Credit
- 1C Credit <-> Counterparty Debit
- Multi-tier matching (Exact, Date-Shifted, Fuzzy Document, Amount Discrepancy, Missing)
"""
import re
from datetime import date
from typing import List, Optional, Tuple, Set

from .models import (
    ReconciliationItem,
    ReconciliationStatement,
    MatchPair,
    MatchStatus,
    ReconciliationReport
)


class StatementMatcher:
    def __init__(self, date_tolerance_days: int = 5, mirror_sides: bool = True):
        """
        :param date_tolerance_days: допустимое расхождение дат (в днях) для одного платежа/отгрузки
        :param mirror_sides: True для стандартного зеркального правила бухучета РФ (Дт 1С == Кт контрагента)
        """
        self.date_tolerance_days = date_tolerance_days
        self.mirror_sides = mirror_sides

    def reconcile(
        self, 
        statement_1c: ReconciliationStatement, 
        statement_cp: ReconciliationStatement
    ) -> ReconciliationReport:
        items_1c = list(statement_1c.items)
        items_cp = list(statement_cp.items)

        matched_pairs: List[MatchPair] = []
        used_cp_ids: Set[str] = set()
        used_1c_ids: Set[str] = set()

        pair_counter = 1

        # =====================================================================
        # ТИР 1: Точное совпадение (Сумма + Номер документа + Дата в допуске)
        # =====================================================================
        for it_1c in items_1c:
            if it_1c.id in used_1c_ids:
                continue

            best_cp = None
            best_days_diff = 999

            for it_cp in items_cp:
                if it_cp.id in used_cp_ids:
                    continue

                if self._amounts_match(it_1c, it_cp):
                    if self._doc_numbers_match(it_1c.doc_number, it_cp.doc_number):
                        days_diff = self._calc_date_diff(it_1c.date, it_cp.date)
                        if days_diff is not None and days_diff <= self.date_tolerance_days:
                            if days_diff < best_days_diff:
                                best_days_diff = days_diff
                                best_cp = it_cp

            if best_cp:
                used_1c_ids.add(it_1c.id)
                used_cp_ids.add(best_cp.id)
                it_1c.matched = True
                best_cp.matched = True
                matched_pairs.append(MatchPair(
                    id=f"pair_{pair_counter}",
                    status=MatchStatus.EXACT,
                    item_1c=it_1c,
                    item_counterparty=best_cp,
                    amount_1c=it_1c.amount,
                    amount_counterparty=best_cp.amount,
                    diff_amount=0.0,
                    date_diff_days=best_days_diff if best_days_diff != 999 else None,
                    explanation=f"Полное совпадение (№ {it_1c.doc_number}, разница дат: {best_days_diff} дн.)"
                ))
                pair_counter += 1

        # =====================================================================
        # ТИР 2: Сумма + Номер документа (Дата сдвинута больше допуска)
        # =====================================================================
        for it_1c in items_1c:
            if it_1c.id in used_1c_ids:
                continue

            for it_cp in items_cp:
                if it_cp.id in used_cp_ids:
                    continue

                if self._amounts_match(it_1c, it_cp):
                    if self._doc_numbers_match(it_1c.doc_number, it_cp.doc_number):
                        days_diff = self._calc_date_diff(it_1c.date, it_cp.date)
                        used_1c_ids.add(it_1c.id)
                        used_cp_ids.add(it_cp.id)
                        it_1c.matched = True
                        it_cp.matched = True
                        matched_pairs.append(MatchPair(
                            id=f"pair_{pair_counter}",
                            status=MatchStatus.AMOUNT_AND_DOC,
                            item_1c=it_1c,
                            item_counterparty=it_cp,
                            amount_1c=it_1c.amount,
                            amount_counterparty=it_cp.amount,
                            diff_amount=0.0,
                            date_diff_days=days_diff,
                            explanation=f"Совпали сумма и номер документа (сдвиг даты: {days_diff or 'н/д'} дн.)"
                        ))
                        pair_counter += 1
                        break

        # =====================================================================
        # ТИР 3: Сумма + Дата (Номер документа отличается или отсутствует)
        # =====================================================================
        for it_1c in items_1c:
            if it_1c.id in used_1c_ids:
                continue

            best_cp = None
            best_days_diff = 999

            for it_cp in items_cp:
                if it_cp.id in used_cp_ids:
                    continue

                if self._amounts_match(it_1c, it_cp):
                    days_diff = self._calc_date_diff(it_1c.date, it_cp.date)
                    if days_diff is not None and days_diff <= self.date_tolerance_days:
                        if days_diff < best_days_diff:
                            best_days_diff = days_diff
                            best_cp = it_cp

            if best_cp:
                used_1c_ids.add(it_1c.id)
                used_cp_ids.add(best_cp.id)
                it_1c.matched = True
                best_cp.matched = True
                matched_pairs.append(MatchPair(
                    id=f"pair_{pair_counter}",
                    status=MatchStatus.AMOUNT_AND_DATE,
                    item_1c=it_1c,
                    item_counterparty=best_cp,
                    amount_1c=it_1c.amount,
                    amount_counterparty=best_cp.amount,
                    diff_amount=0.0,
                    date_diff_days=best_days_diff,
                    explanation=f"Сумма совпала по дате (разные номера: 1C='{it_1c.doc_number}', Контрагент='{best_cp.doc_number}')"
                ))
                pair_counter += 1

        # =====================================================================
        # ТИР 4: Номер документа совпал, но сумма РАЗЛИЧАЕТСЯ (Расхождение суммы)
        # =====================================================================
        for it_1c in items_1c:
            if it_1c.id in used_1c_ids:
                continue

            if not it_1c.doc_number or len(it_1c.doc_number) < 2:
                continue

            for it_cp in items_cp:
                if it_cp.id in used_cp_ids:
                    continue

                # Должно быть совпадение стороны операции (Дт 1С <-> Кт КП или наоборот)
                if self._same_accounting_side(it_1c, it_cp):
                    if self._doc_numbers_match(it_1c.doc_number, it_cp.doc_number):
                        days_diff = self._calc_date_diff(it_1c.date, it_cp.date)
                        if days_diff is None or days_diff <= 14:
                            used_1c_ids.add(it_1c.id)
                            used_cp_ids.add(it_cp.id)
                            diff_val = round(it_1c.amount - it_cp.amount, 2)
                            matched_pairs.append(MatchPair(
                                id=f"pair_{pair_counter}",
                                status=MatchStatus.AMOUNT_DIFF,
                                item_1c=it_1c,
                                item_counterparty=it_cp,
                                amount_1c=it_1c.amount,
                                amount_counterparty=it_cp.amount,
                                diff_amount=diff_val,
                                date_diff_days=days_diff,
                                explanation=f"РАСХОЖДЕНИЕ В СУММЕ: разница {diff_val:+.2f} руб. (№ документа '{it_1c.doc_number}')"
                            ))
                            pair_counter += 1
                            break

        # =====================================================================
        # ТИР 5: Несопоставленные строки
        # =====================================================================
        # 1C только
        for it_1c in items_1c:
            if it_1c.id not in used_1c_ids:
                side_str = "Дебет (отгрузка/услуга)" if it_1c.is_debit else "Кредит (оплата/поступление)"
                matched_pairs.append(MatchPair(
                    id=f"pair_{pair_counter}",
                    status=MatchStatus.ONLY_1C,
                    item_1c=it_1c,
                    item_counterparty=None,
                    amount_1c=it_1c.amount,
                    amount_counterparty=0.0,
                    diff_amount=it_1c.amount,
                    date_diff_days=None,
                    explanation=f"Документ есть в 1С, но ОТСУТСТВУЕТ у контрагента ({side_str})"
                ))
                pair_counter += 1

        # Контрагент только
        for it_cp in items_cp:
            if it_cp.id not in used_cp_ids:
                side_str = "Дебет (оплата)" if it_cp.is_debit else "Кредит (отгрузка)"
                matched_pairs.append(MatchPair(
                    id=f"pair_{pair_counter}",
                    status=MatchStatus.ONLY_COUNTERPARTY,
                    item_1c=None,
                    item_counterparty=it_cp,
                    amount_1c=0.0,
                    amount_counterparty=it_cp.amount,
                    diff_amount=-it_cp.amount,
                    date_diff_days=None,
                    explanation=f"Документ есть у контрагента, но ОТСУТСТВУЕТ в 1С ({side_str})"
                ))
                pair_counter += 1

        # =====================================================================
        # Расчет итогов и метрик
        # =====================================================================
        matched_count = sum(1 for p in matched_pairs if p.is_matched)
        discrepancy_count = len(matched_pairs) - matched_count

        # Сравнение сальдо и оборотов с учетом зеркальности:
        # Дебет 1С должен быть равен Кредиту Контрагента
        # Кредит 1С должен быть равен Дебету Контрагента
        if self.mirror_sides:
            diff_init = round(statement_1c.initial_balance - (-statement_cp.initial_balance if statement_cp.initial_balance else 0.0), 2)
            diff_dt = round(statement_1c.turnover_debit - statement_cp.turnover_credit, 2)
            diff_kt = round(statement_1c.turnover_credit - statement_cp.turnover_debit, 2)
            diff_final = round(statement_1c.final_balance - (-statement_cp.final_balance if statement_cp.final_balance else 0.0), 2)
        else:
            diff_init = round(statement_1c.initial_balance - statement_cp.initial_balance, 2)
            diff_dt = round(statement_1c.turnover_debit - statement_cp.turnover_debit, 2)
            diff_kt = round(statement_1c.turnover_credit - statement_cp.turnover_credit, 2)
            diff_final = round(statement_1c.final_balance - statement_cp.final_balance, 2)

        has_discrepancies = (
            discrepancy_count > 0 or 
            abs(diff_init) > 0.01 or 
            abs(diff_dt) > 0.01 or 
            abs(diff_kt) > 0.01 or 
            abs(diff_final) > 0.01
        )

        if not has_discrepancies:
            verdict = "СВЕРКА ПРОЙДЕНА УСПЕШНО: Расхождений между 1С и актом контрагента не обнаружено."
        else:
            verdict = f"ОБНАРУЖЕНЫ РАСХОЖДЕНИЯ: не сошлось {discrepancy_count} операций. "
            if abs(diff_final) > 0.01:
                verdict += f"Разница конечного сальдо: {diff_final:+.2f} руб."

        return ReconciliationReport(
            statement_1c=statement_1c,
            statement_counterparty=statement_cp,
            pairs=matched_pairs,
            total_items_1c=len(items_1c),
            total_items_cp=len(items_cp),
            matched_count=matched_count,
            discrepancy_count=discrepancy_count,
            diff_initial_balance=diff_init,
            diff_turnover_debit=diff_dt,
            diff_turnover_credit=diff_kt,
            diff_final_balance=diff_final,
            has_discrepancies=has_discrepancies,
            summary_verdict=verdict
        )

    def _same_accounting_side(self, item_1c: ReconciliationItem, item_cp: ReconciliationItem) -> bool:
        """Проверяет соответствие сторон учета (зеркальность)."""
        if self.mirror_sides:
            return (item_1c.is_debit and not item_cp.is_debit) or (not item_1c.is_debit and item_cp.is_debit)
        return item_1c.is_debit == item_cp.is_debit

    def _amounts_match(self, item_1c: ReconciliationItem, item_cp: ReconciliationItem) -> bool:
        """Проверяет равенство сумм с учетом зеркальности учета."""
        if not self._same_accounting_side(item_1c, item_cp):
            return False
        return abs(item_1c.amount - item_cp.amount) < 0.01

    def _doc_numbers_match(self, num1: str, num2: str) -> bool:
        if not num1 or not num2:
            return False
        
        # Очистка от лидирующих нулей, пробелов, дефисов
        n1 = re.sub(r'^[0\s\-]+', '', num1).strip().lower()
        n2 = re.sub(r'^[0\s\-]+', '', num2).strip().lower()

        if not n1 or not n2:
            return False

        if n1 == n2:
            return True

        # Если один номер входит в другой как целая часть (например "0000-000123" и "123")
        if (len(n1) >= 2 and n1 in n2) or (len(n2) >= 2 and n2 in n1):
            return True

        return False

    def _calc_date_diff(self, d1: Optional[date], d2: Optional[date]) -> Optional[int]:
        if not d1 or not d2:
            return None
        return abs((d1 - d2).days)
