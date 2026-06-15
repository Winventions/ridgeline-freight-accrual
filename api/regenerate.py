import cgi
import io
import json
import sys
import tempfile
import traceback
import zipfile
from http.server import BaseHTTPRequestHandler
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from run_accrual import build_outputs  # noqa: E402


REQUIRED_FILES = [
    "shipments_apr2026.csv",
    "rate_card_peak_logistics.csv",
    "rate_card_heartland_freight.csv",
    "rate_card_coastal_express.csv",
    "freight_invoices_oct2025_mar2026_v2.csv",
    "denise_accruals_v2.csv",
]


class handler(BaseHTTPRequestHandler):
    def _send_json(self, status, payload):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        self._send_json(
            200,
            {
                "status": "ready",
                "required_files": REQUIRED_FILES,
                "message": "POST multipart/form-data with each required CSV file.",
            },
        )

    def do_POST(self):
        content_type = self.headers.get("content-type", "")
        if "multipart/form-data" not in content_type:
            self._send_json(400, {"error": "Expected multipart/form-data upload."})
            return

        try:
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": content_type,
                    "CONTENT_LENGTH": self.headers.get("content-length", "0"),
                },
            )

            with tempfile.TemporaryDirectory() as tmp:
                tmp_path = Path(tmp)
                data_dir = tmp_path / "data"
                output_dir = tmp_path / "output"
                data_dir.mkdir()
                output_dir.mkdir()

                missing = []
                for file_name in REQUIRED_FILES:
                    field = form[file_name] if file_name in form else None
                    if field is None or not getattr(field, "file", None):
                        missing.append(file_name)
                        continue
                    content = field.file.read()
                    if not content:
                        missing.append(file_name)
                        continue
                    (data_dir / file_name).write_bytes(content)

                if missing:
                    self._send_json(400, {"error": "Missing required files.", "missing": missing})
                    return

                workbook, journal, exceptions, tie_out = build_outputs(data_dir, output_dir)

                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as archive:
                    for path in [workbook, journal, exceptions]:
                        archive.write(path, arcname=path.name)
                    summary = {
                        "tie_out": tie_out,
                        "generated_files": [workbook.name, journal.name, exceptions.name],
                    }
                    archive.writestr("run_summary.json", json.dumps(summary, indent=2))

                body = zip_buffer.getvalue()
                self.send_response(200)
                self.send_header("Content-Type", "application/zip")
                self.send_header(
                    "Content-Disposition",
                    'attachment; filename="ridgeline_freight_accrual_outputs.zip"',
                )
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
        except Exception as exc:
            self._send_json(
                500,
                {
                    "error": str(exc),
                    "trace": traceback.format_exc(limit=8),
                },
            )
