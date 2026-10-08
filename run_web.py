# -*- coding: utf-8 -*-
"""
Web server launcher for the Accounting Reconciliation application.
Zero external server dependencies (uses Python standard library http.server).
"""
import os
import sys
import json
import uuid
import shutil
from pathlib import Path
from http.server import HTTPServer, SimpleHTTPRequestHandler
import urllib.parse

sys.stdout.reconfigure(encoding='utf-8')

from core.models import ReconciliationReport, MatchPair, MatchStatus, ReconciliationItem
from core.excel_parser import Excel1CParser
from core.pdf_parser import PDFCounterpartyParser
from core.matcher import StatementMatcher
from core.exporter import ExcelReportExporter
from core.qwen_client import QwenVLClient

BASE_DIR = Path(__file__).resolve().parent
WEB_DIR = BASE_DIR / "web"
UPLOADS_DIR = BASE_DIR / "uploads"
REPORTS_DIR = BASE_DIR / "reports"
SAMPLES_DIR = BASE_DIR / "samples"

UPLOADS_DIR.mkdir(exist_ok=True)
REPORTS_DIR.mkdir(exist_ok=True)


def report_to_dict(report: ReconciliationReport, report_filename: str) -> dict:
    def item_to_dict(item: ReconciliationItem):
        if not item:
            return None
        return {
            "id": item.id,
            "source": item.source,
            "date_str": item.date_str,
            "doc_name": item.doc_name,
            "doc_number": item.doc_number,
            "doc_raw": item.doc_raw,
            "debit": item.debit,
            "credit": item.credit,
            "amount": item.amount,
            "is_debit": item.is_debit
        }

    pairs_list = []
    for p in report.pairs:
        pairs_list.append({
            "id": p.id,
            "status": p.status.value,
            "item_1c": item_to_dict(p.item_1c),
            "item_cp": item_to_dict(p.item_counterparty),
            "amount_1c": p.amount_1c,
            "amount_cp": p.amount_counterparty,
            "diff_amount": p.diff_amount,
            "date_diff_days": p.date_diff_days,
            "explanation": p.explanation
        })

    return {
        "report_file": report_filename,
        "has_discrepancies": report.has_discrepancies,
        "summary_verdict": report.summary_verdict,
        "total_items_1c": report.total_items_1c,
        "total_items_cp": report.total_items_cp,
        "matched_count": report.matched_count,
        "discrepancy_count": report.discrepancy_count,
        "diff_initial_balance": report.diff_initial_balance,
        "diff_turnover_debit": report.diff_turnover_debit,
        "diff_turnover_credit": report.diff_turnover_credit,
        "diff_final_balance": report.diff_final_balance,
        "statement_1c": {
            "organization": report.statement_1c.organization,
            "counterparty": report.statement_1c.counterparty,
            "period_str": report.statement_1c.period_str,
            "initial_balance": report.statement_1c.initial_balance,
            "turnover_debit": report.statement_1c.turnover_debit,
            "turnover_credit": report.statement_1c.turnover_credit,
            "final_balance": report.statement_1c.final_balance
        },
        "statement_cp": {
            "organization": report.statement_counterparty.organization,
            "counterparty": report.statement_counterparty.counterparty,
            "period_str": report.statement_counterparty.period_str,
            "initial_balance": report.statement_counterparty.initial_balance,
            "turnover_debit": report.statement_counterparty.turnover_debit,
            "turnover_credit": report.statement_counterparty.turnover_credit,
            "final_balance": report.statement_counterparty.final_balance
        },
        "pairs": pairs_list
    }


class ReconcileHTTPHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(WEB_DIR), **kwargs)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == "/api/reconcile-demo":
            self.handle_demo()
            return

        if path.startswith("/api/download/"):
            filename = os.path.basename(path)
            file_path = REPORTS_DIR / filename
            if file_path.exists():
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Content-Length", str(file_path.stat().st_size))
                self.end_headers()
                with open(file_path, "rb") as f:
                    self.wfile.write(f.read())
            else:
                self.send_error(404, "Report file not found")
            return

        # Раздача статических файлов web/
        return super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/reconcile":
            self.handle_reconcile()
            return
        self.send_error(404, "Not Found")

    def handle_demo(self):
        try:
            excel_path = SAMPLES_DIR / "act_1c_sample.xlsx"
            pdf_path = SAMPLES_DIR / "act_counterparty_sample.pdf"

            if not excel_path.exists() or not pdf_path.exists():
                from samples.generate_samples import generate_1c_excel, generate_counterparty_scan
                SAMPLES_DIR.mkdir(exist_ok=True)
                generate_1c_excel(str(excel_path))
                generate_counterparty_scan(
                    str(SAMPLES_DIR / "act_counterparty_scan.png"),
                    str(pdf_path)
                )

            p1 = Excel1CParser()
            s_1c = p1.parse(str(excel_path))

            p_cp = PDFCounterpartyParser()
            s_cp = p_cp.parse(str(pdf_path), force_ocr=True)

            matcher = StatementMatcher(date_tolerance_days=5)
            report = matcher.reconcile(s_1c, s_cp)

            rep_name = f"Протокол_сверки_ДЕМО_{uuid.uuid4().hex[:6]}.xlsx"
            out_file = REPORTS_DIR / rep_name
            exporter = ExcelReportExporter()
            exporter.export(report, str(out_file))

            resp_dict = report_to_dict(report, rep_name)
            self._send_json(resp_dict)
        except Exception as e:
            self._send_json({"error": str(e)}, status=500)

    def handle_reconcile(self):
        try:
            content_type = self.headers.get("Content-Type", "")
            if not content_type.startswith("multipart/form-data"):
                self._send_json({"error": "Content-Type must be multipart/form-data"}, status=400)
                return

            boundary = content_type.split("boundary=")[1].encode()
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length)

            parts = body.split(b"--" + boundary)
            files = {}
            fields = {}

            for part in parts:
                if not part or part == b"--\r\n" or part == b"--":
                    continue
                header_sep = part.find(b"\r\n\r\n")
                if header_sep == -1:
                    continue
                header_data = part[:header_sep].decode("utf-8", errors="replace")
                content = part[header_sep + 4:].rstrip(b"\r\n")

                if 'filename="' in header_data:
                    # Извлечение имени поля и файла
                    field_name = ""
                    filename = ""
                    for line in header_data.splitlines():
                        if "Content-Disposition:" in line:
                            for item in line.split(";"):
                                item = item.strip()
                                if item.startswith("name="):
                                    field_name = item.split("=")[1].strip('"')
                                elif item.startswith("filename="):
                                    filename = item.split("=")[1].strip('"')
                    if field_name and filename:
                        files[field_name] = (filename, content)
                else:
                    # Текстовое поле
                    for line in header_data.splitlines():
                        if "Content-Disposition:" in line:
                            for item in line.split(";"):
                                item = item.strip()
                                if item.startswith("name="):
                                    field_name = item.split("=")[1].strip('"')
                                    fields[field_name] = content.decode("utf-8", errors="replace").strip()

            if "file_1c" not in files or "file_cp" not in files:
                self._send_json({"error": "Требуются оба файла: file_1c и file_cp"}, status=400)
                return

            req_id = uuid.uuid4().hex[:8]
            f1c_name, f1c_bytes = files["file_1c"]
            fcp_name, fcp_bytes = files["file_cp"]

            saved_1c = UPLOADS_DIR / f"{req_id}_{f1c_name}"
            saved_cp = UPLOADS_DIR / f"{req_id}_{fcp_name}"

            with open(saved_1c, "wb") as f:
                f.write(f1c_bytes)
            with open(saved_cp, "wb") as f:
                f.write(fcp_bytes)

            model = fields.get("model", "qwen3-vl:4b-instruct")
            tolerance = int(fields.get("tolerance", 5))
            force_ocr = fields.get("force_ocr", "true").lower() in ("true", "1", "yes")

            # Выполнение парсинга и сверки
            p1 = Excel1CParser()
            s_1c = p1.parse(str(saved_1c))

            p_cp = PDFCounterpartyParser(model_name=model)
            s_cp = p_cp.parse(str(saved_cp), force_ocr=force_ocr)

            matcher = StatementMatcher(date_tolerance_days=tolerance)
            report = matcher.reconcile(s_1c, s_cp)

            rep_name = f"Протокол_сверки_{req_id}.xlsx"
            out_file = REPORTS_DIR / rep_name
            exporter = ExcelReportExporter()
            exporter.export(report, str(out_file))

            resp_dict = report_to_dict(report, rep_name)
            self._send_json(resp_dict)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self._send_json({"error": f"Ошибка обработки: {str(e)}"}, status=500)

    def _send_json(self, data: dict, status: int = 200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def run_server(port: int = 8080):
    server_address = ('', port)
    httpd = HTTPServer(server_address, ReconcileHTTPHandler)
    print("=" * 60)
    print(f"🚀 Веб-приложение сверки актов запущено!")
    print(f"👉 Откройте в браузере: http://127.0.0.1:{port}")
    print(f"⚡ Модель: Qwen3-VL (Ollama) на AMD Radeon RX 7600")
    print("=" * 60)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nСервер остановлен.")


if __name__ == "__main__":
    port = 8080
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            pass
    run_server(port)
