# -*- coding: utf-8 -*-
"""
Command-line interface (CLI) for accounting reconciliation.
"""
import sys
import argparse
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

from rich.console import Console
from rich.table import Table
from rich.panel import Panel

from core.excel_parser import Excel1CParser
from core.pdf_parser import PDFCounterpartyParser
from core.matcher import StatementMatcher
from core.exporter import ExcelReportExporter
from core.models import MatchStatus

console = Console()


def cmd_reconcile(args):
    excel_path = Path(args.excel)
    pdf_path = Path(args.pdf)

    if not excel_path.exists():
        console.print(f"[bold red]Ошибка:[/bold red] Файл 1С не найден: {excel_path}")
        sys.exit(1)
    if not pdf_path.exists():
        console.print(f"[bold red]Ошибка:[/bold red] Файл контрагента не найден: {pdf_path}")
        sys.exit(1)

    console.print(Panel(f"[bold cyan]Сверка актов взаиморасчетов[/bold cyan]\n1С: {excel_path.name}\nКонтрагент: {pdf_path.name}\nМодель: {args.model}", title="Параметры запуска"))

    with console.status("[bold green]Парсинг выгрузки 1С..."):
        p_1c = Excel1CParser()
        stmt_1c = p_1c.parse(str(excel_path))
    console.print(f"✔ 1С: найдено [bold]{len(stmt_1c.items)}[/bold] операций (Период: {stmt_1c.period_str or 'не указан'})")

    with console.status(f"[bold green]Распознавание акта контрагента через {args.model}..."):
        p_cp = PDFCounterpartyParser(model_name=args.model)
        stmt_cp = p_cp.parse(str(pdf_path), force_ocr=args.force_ocr)
    console.print(f"✔ Контрагент: найдено [bold]{len(stmt_cp.items)}[/bold] операций")

    with console.status("[bold green]Сопоставление операций и расчет сальдо..."):
        matcher = StatementMatcher(date_tolerance_days=args.tolerance)
        report = matcher.reconcile(stmt_1c, stmt_cp)

    # Итоги
    if report.has_discrepancies:
        console.print(Panel(f"[bold red]{report.summary_verdict}[/bold red]", title="Результат сверки", border_style="red"))
    else:
        console.print(Panel(f"[bold green]{report.summary_verdict}[/bold green]", title="Результат сверки", border_style="green"))

    # Таблица сальдо
    bal_table = Table(title="Сравнение начального/конечного сальдо и оборотов", show_header=True)
    bal_table.add_column("Показатель", style="bold")
    bal_table.add_column("По данным 1С", justify="right")
    bal_table.add_column("По данным Контрагента", justify="right")
    bal_table.add_column("Разница", justify="right")

    bal_table.add_row(
        "Сальдо начальное", 
        f"{report.statement_1c.initial_balance:,.2f} ₽", 
        f"{report.statement_counterparty.initial_balance:,.2f} ₽", 
        f"{report.diff_initial_balance:+,.2f} ₽"
    )
    bal_table.add_row(
        "Обороты Дебет 1С / Кредит КП", 
        f"{report.statement_1c.turnover_debit:,.2f} ₽", 
        f"{report.statement_counterparty.turnover_credit:,.2f} ₽", 
        f"{report.diff_turnover_debit:+,.2f} ₽"
    )
    bal_table.add_row(
        "Обороты Кредит 1С / Дебет КП", 
        f"{report.statement_1c.turnover_credit:,.2f} ₽", 
        f"{report.statement_counterparty.turnover_debit:,.2f} ₽", 
        f"{report.diff_turnover_credit:+,.2f} ₽"
    )
    bal_table.add_row(
        "Сальдо конечное", 
        f"{report.statement_1c.final_balance:,.2f} ₽", 
        f"{report.statement_counterparty.final_balance:,.2f} ₽", 
        f"{report.diff_final_balance:+,.2f} ₽"
    )
    console.print(bal_table)

    # Детальная таблица расхождений
    discrepancies = [p for p in report.pairs if not p.is_matched]
    if discrepancies:
        d_table = Table(title="Ведомость обнаруженных расхождений", show_header=True)
        d_table.add_column("Статус", style="bold red")
        d_table.add_column("Разница", justify="right")
        d_table.add_column("Дата 1С")
        d_table.add_column("Документ 1С")
        d_table.add_column("Сумма 1С", justify="right")
        d_table.add_column("Дата КП")
        d_table.add_column("Документ КП")
        d_table.add_column("Сумма КП", justify="right")
        d_table.add_column("Пояснение")

        for p in discrepancies:
            d_table.add_row(
                p.status.value,
                f"{p.diff_amount:+,.2f} ₽",
                p.item_1c.date_str if p.item_1c else "—",
                p.item_1c.doc_raw if p.item_1c else "—",
                f"{p.amount_1c:,.2f} ₽" if p.amount_1c else "—",
                p.item_counterparty.date_str if p.item_counterparty else "—",
                p.item_counterparty.doc_raw if p.item_counterparty else "—",
                f"{p.amount_counterparty:,.2f} ₽" if p.amount_counterparty else "—",
                p.explanation
            )
        console.print(d_table)

    # Сохранение отчета
    output_path = args.output or f"Протокол_сверки_{excel_path.stem}.xlsx"
    exporter = ExcelReportExporter()
    exporter.export(report, output_path)
    console.print(f"\n[bold green]✔ Полный протокол сверки сохранен в файл:[/bold green] [bold underline]{output_path}[/bold underline]")


def main():
    parser = argparse.ArgumentParser(description="Утилита сверки актов взаиморасчетов бухгалтерии (1C + Scans/PDF)")
    subparsers = parser.add_subparsers(dest="command")

    # Команда reconcile
    rec_parser = subparsers.add_parser("reconcile", help="Выполнить сверку между 1C Excel и актом контрагента")
    rec_parser.add_argument("--excel", "-e", required=True, help="Путь к акту 1С (.xlsx, .xls)")
    rec_parser.add_argument("--pdf", "-p", required=True, help="Путь к акту контрагента (PDF, скан, изображение)")
    rec_parser.add_argument("--output", "-o", default=None, help="Путь для сохранения Excel отчета")
    rec_parser.add_argument("--model", "-m", default="qwen3-vl:4b-instruct", help="Модель Ollama (по умолчанию: qwen3-vl:4b-instruct)")
    rec_parser.add_argument("--tolerance", "-t", type=int, default=5, help="Допуск дат в днях (по умолчанию: 5)")
    rec_parser.add_argument("--force-ocr", action="store_true", default=True, help="Принудительный OCR скана")

    args = parser.parse_args()
    if args.command == "reconcile":
        cmd_reconcile(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
