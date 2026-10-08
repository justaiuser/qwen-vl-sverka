# -*- coding: utf-8 -*-
"""
End-to-end reconciliation test script.
"""
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

from core.excel_parser import Excel1CParser
from core.pdf_parser import PDFCounterpartyParser
from core.matcher import StatementMatcher
from core.exporter import ExcelReportExporter

excel_path = os.path.join("samples", "act_1c_sample.xlsx")
pdf_path = os.path.join("samples", "act_counterparty_sample.pdf")
report_path = os.path.join("samples", "report_test.xlsx")

print("1. Парсинг 1С Excel...")
p1 = Excel1CParser()
s_1c = p1.parse(excel_path)
print(f"   Организация: {s_1c.organization}, Контрагент: {s_1c.counterparty}")
print(f"   Период: {s_1c.period_str}")
print(f"   Сальдо нач: {s_1c.initial_balance}, Обороты Дт: {s_1c.turnover_debit}, Кт: {s_1c.turnover_credit}")
print(f"   Найдено строк 1С: {len(s_1c.items)}")
for it in s_1c.items:
    print(f"     * {it.date_str} | {it.doc_raw} | Дт: {it.debit} | Кт: {it.credit}")

print("\n2. Парсинг акта контрагента (скан PDF через Qwen-VL)...")
p_cp = PDFCounterpartyParser()
s_cp = p_cp.parse(pdf_path, force_ocr=True)
print(f"   Организация: {s_cp.organization}, Контрагент: {s_cp.counterparty}")
print(f"   Период: {s_cp.period_str}")
print(f"   Сальдо нач: {s_cp.initial_balance}, Обороты Дт: {s_cp.turnover_debit}, Кт: {s_cp.turnover_credit}")
print(f"   Найдено строк контрагента: {len(s_cp.items)}")
for it in s_cp.items:
    print(f"     * {it.date_str} | {it.doc_raw} | Дт: {it.debit} | Кт: {it.credit}")

print("\n3. Сопоставление операций...")
matcher = StatementMatcher(date_tolerance_days=5)
report = matcher.reconcile(s_1c, s_cp)

print("\n=== РЕЗУЛЬТАТЫ СВЕРКИ ===")
print(f"Вердикт: {report.summary_verdict}")
print(f"Всего 1С: {report.total_items_1c} | Всего КП: {report.total_items_cp}")
print(f"Сверено: {report.matched_count} | Расхождений: {report.discrepancy_count}")
print(f"Разница сальдо начального: {report.diff_initial_balance}")
print(f"Разница оборотов Дт: {report.diff_turnover_debit} | Кт: {report.diff_turnover_credit}")
print(f"Разница сальдо конечного: {report.diff_final_balance}")

print("\nДетализация пар:")
for p in report.pairs:
    status_icon = "✅" if p.is_matched else "❌"
    print(f" {status_icon} [{p.status.value}] Разница: {p.diff_amount:+.2f} | 1С: {p.item_1c.doc_raw if p.item_1c else '---'} ({p.amount_1c}) <-> КП: {p.item_counterparty.doc_raw if p.item_counterparty else '---'} ({p.amount_counterparty})")
    print(f"    Пояснение: {p.explanation}")

print("\n4. Экспорт отчета в Excel...")
exporter = ExcelReportExporter()
exporter.export(report, report_path)
print(f"Отчет сохранен в: {report_path}")
