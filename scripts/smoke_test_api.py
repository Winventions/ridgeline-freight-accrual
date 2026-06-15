import mimetypes
import sys
import threading
import urllib.request
from http.server import HTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from api.regenerate import REQUIRED_FILES, handler  # noqa: E402


def multipart_body(files):
    boundary = "----ridgeline-smoke-test"
    chunks = []
    for file_name in files:
        path = ROOT / "data" / file_name
        content_type = mimetypes.guess_type(path.name)[0] or "text/csv"
        chunks.append(f"--{boundary}\r\n".encode("utf-8"))
        chunks.append(
            (
                f'Content-Disposition: form-data; name="{file_name}"; filename="{file_name}"\r\n'
                f"Content-Type: {content_type}\r\n\r\n"
            ).encode("utf-8")
        )
        chunks.append(path.read_bytes())
        chunks.append(b"\r\n")
    chunks.append(f"--{boundary}--\r\n".encode("utf-8"))
    return boundary, b"".join(chunks)


def main():
    server = HTTPServer(("127.0.0.1", 0), handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        base_url = f"http://127.0.0.1:{server.server_port}"
        with urllib.request.urlopen(f"{base_url}/", timeout=20) as response:
            home = response.read()
            if response.status != 200 or b"Ridgeline Foods Freight Accrual Engine" not in home:
                raise RuntimeError("Homepage smoke test failed")

        with urllib.request.urlopen(f"{base_url}/api/regenerate", timeout=20) as response:
            api_health = response.read()
            if response.status != 200 or b"required_files" not in api_health:
                raise RuntimeError("API health smoke test failed")

        boundary, body = multipart_body(REQUIRED_FILES)
        request = urllib.request.Request(
            f"{base_url}/api/regenerate",
            data=body,
            method="POST",
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        )
        with urllib.request.urlopen(request, timeout=60) as response:
            payload = response.read()
            content_type = response.headers.get("Content-Type", "")
            if response.status != 200 or content_type != "application/zip" or not payload.startswith(b"PK"):
                raise RuntimeError(f"Unexpected response: {response.status} {content_type}")
            print(f"API smoke test passed: {len(payload)} bytes")
    finally:
        server.shutdown()


if __name__ == "__main__":
    main()
