"""Loopback-only application. No cloud, API keys or additional web dependencies."""
from __future__ import annotations

import argparse
import json
import os
import secrets
import socket
import subprocess
import sys
import threading
import uuid
import webbrowser
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit

from data import build_package, export_zip

ROOT = Path(__file__).resolve().parent
RUNS = ROOT / "runs"
TOKEN = secrets.token_urlsafe(32)
JOBS = {}
LOCK = threading.Lock()
COLLECT_LOCK = threading.Lock()


class LocalServer(ThreadingHTTPServer):
    allow_reuse_address = False

    def server_bind(self):
        if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()


def status(run, **values):
    with LOCK:
        JOBS.setdefault(run, {}).update(values)
        state = dict(JOBS[run])
        target = RUNS / run / "status.json"
        tmp = target.with_suffix(".tmp")
        tmp.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
        tmp.replace(target)


def work(run, files):
    try:
        status(run, phase="working", message="Читаю файлы…")
        package = build_package(files["selection"], files["abc"], files.get("competitors"), files.get("queries"),
                                lambda message: status(run, message=message))
        folder = RUNS / run
        names = json.loads((folder / "original_names.json").read_text(encoding="utf-8"))
        for source in package["sources"]:
            source["original_file"] = names.get(source["file"], source["file"])
        (folder / "data.json").write_text(json.dumps(package, ensure_ascii=False, indent=2), encoding="utf-8")
        export_zip(package, folder / "analysis.zip")
        status(run, phase="done", message="Пакет готов", summary=package["summary"], has_queries="queries" in files)
    except Exception as exc:
        status(run, phase="error", message=str(exc))


def collect(run):
    try:
        status(run, phase="collecting", message="Откроется Chrome. Войдите в MPSTATS; сбор продолжится автоматически.")
        folder = RUNS / run
        with (folder / "collector.log").open("w", encoding="utf-8") as log:
            child = subprocess.run([sys.executable, "-u", str(ROOT / "worker.py"), str(folder)],
                                   cwd=ROOT.parent, stdout=log, stderr=subprocess.STDOUT,
                                   env={**os.environ, "PYTHONIOENCODING": "utf-8"})
        if child.returncode:
            raise RuntimeError("Сбор MPSTATS прерван. Подробности в журнале; повторный запуск использует кэш.")
        config = json.loads((folder / "files.json").read_text(encoding="utf-8"))
        config["competitors"] = "competitors.xlsx"
        (folder / "files.json").write_text(json.dumps(config), encoding="utf-8")
        work(run, {key: folder / value for key, value in config.items()})
    except Exception as exc:
        status(run, phase="error", message=str(exc))
    finally:
        COLLECT_LOCK.release()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *_):
        pass

    def allowed(self):
        host = self.headers.get("Host", "")
        return host == f"127.0.0.1:{self.server.server_port}"

    def send(self, code, data, content_type="application/json; charset=utf-8", filename=None):
        if isinstance(data, (dict, list)):
            data = json.dumps(data, ensure_ascii=False).encode("utf-8")
        elif isinstance(data, str):
            data = data.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Content-Security-Policy", "default-src 'self'; script-src 'self'; style-src 'self'; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'")
        if filename:
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if not self.allowed():
            return self.send(403, {"error": "Недопустимый адрес"})
        path = urlsplit(self.path).path
        assets = {"/": ("index.html", "text/html; charset=utf-8"), "/app.js": ("app.js", "text/javascript; charset=utf-8"), "/style.css": ("style.css", "text/css; charset=utf-8")}
        if path in assets:
            name, mime = assets[path]
            return self.send(200, (ROOT / "static" / name).read_bytes(), mime)
        if path == "/api/session":
            return self.send(200, {"token": TOKEN})
        if path == "/api/runs":
            with LOCK:
                result = [{"id": run, **value} for run, value in reversed(list(JOBS.items()))]
            return self.send(200, result)
        parts = path.strip("/").split("/")
        if len(parts) == 3 and parts[0] == "runs" and parts[1] in JOBS:
            files = {"data": ("data.json", "application/json; charset=utf-8"),
                     "download": ("analysis.zip", "application/zip"), "log": ("collector.log", "text/plain; charset=utf-8")}
            if parts[2] in files:
                name, mime = files[parts[2]]
                target = RUNS / parts[1] / name
                if target.is_file():
                    return self.send(200, target.read_bytes(), mime, name if parts[2] == "download" else None)
        return self.send(404, {"error": "Не найдено"})

    def do_POST(self):
        if not self.allowed() or self.headers.get("X-App-Token") != TOKEN:
            return self.send(403, {"error": "Обновите страницу и повторите действие"})
        origin = self.headers.get("Origin")
        if origin and origin != f"http://127.0.0.1:{self.server.server_port}":
            return self.send(403, {"error": "Недопустимый источник"})
        path = urlsplit(self.path).path
        parts = path.strip("/").split("/")
        if len(parts) == 3 and parts[0] == "collect" and parts[2] == "start" and parts[1] in JOBS:
            run = parts[1]
            config = json.loads((RUNS / run / "files.json").read_text(encoding="utf-8"))
            if "queries" not in config or JOBS[run]["phase"] in ("working", "collecting"):
                return self.send(400, {"error": "Нужна готовая таблица запросов и завершённый импорт"})
            if not COLLECT_LOCK.acquire(blocking=False):
                return self.send(409, {"error": "Другой сбор MPSTATS уже выполняется"})
            status(run, phase="collecting", message="Подготавливаю сбор MPSTATS…")
            threading.Thread(target=collect, args=(run,), daemon=True).start()
            return self.send(202, {"id": run})
        if path != "/api/import":
            return self.send(404, {"error": "Не найдено"})
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length <= 0 or length > 150 * 1024 * 1024:
                raise ValueError("Размер загрузки должен быть от 1 байта до 150 МБ.")
            content_type = self.headers.get("Content-Type", "")
            if not content_type.startswith("multipart/form-data;"):
                raise ValueError("Ожидается загрузка файлов")
            message = BytesParser(policy=default).parsebytes(f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode() + self.rfile.read(length))
            uploads = {}
            for part in message.iter_parts():
                key = part.get_param("name", header="content-disposition")
                filename = part.get_filename()
                if key not in ("selection", "abc", "queries", "competitors") or not filename:
                    continue
                suffix = Path(filename).suffix.lower()
                if suffix not in ((".txt", ".csv", ".xlsx") if key == "selection" else (".xlsx",)):
                    raise ValueError("Допустимы XLSX, а для списка SKU также CSV и TXT.")
                payload = part.get_payload(decode=True)
                if not payload:
                    raise ValueError("Загружен пустой файл")
                if suffix == ".xlsx":
                    import io, zipfile
                    with zipfile.ZipFile(io.BytesIO(payload)) as archive:
                        if sum(item.file_size for item in archive.infolist()) > 700 * 1024 * 1024:
                            raise ValueError("Распакованная книга слишком велика (лимит 700 МБ)")
                uploads[key] = (suffix, payload, Path(filename).name)
            if not {"selection", "abc"}.issubset(uploads):
                raise ValueError("Выберите список SKU и ABC-книгу")
            run = uuid.uuid4().hex
            folder = RUNS / run
            folder.mkdir(parents=True)
            files = {}
            original_names = {}
            for key, (suffix, payload, filename) in uploads.items():
                files[key] = folder / (key + suffix)
                files[key].write_bytes(payload)
                original_names[key + suffix] = filename
            (folder / "original_names.json").write_text(json.dumps(original_names, ensure_ascii=False), encoding="utf-8")
            (folder / "files.json").write_text(json.dumps({key: value.name for key, value in files.items()}), encoding="utf-8")
            status(run, phase="working", message="Загрузка завершена", title=original_names[files["selection"].name])
            threading.Thread(target=work, args=(run, files), daemon=True).start()
            self.send(202, {"id": run})
        except Exception as exc:
            self.send(400, {"error": str(exc)})


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()
    RUNS.mkdir(exist_ok=True)
    for path in sorted(RUNS.glob("*/status.json"), key=lambda p: p.stat().st_mtime):
        try:
            state = json.loads(path.read_text(encoding="utf-8"))
            if state.get("phase") in ("working", "collecting"):
                state.update(phase="error", message="Предыдущий запуск был прерван. Повторите импорт или сбор.")
            JOBS[path.parent.name] = state
        except (ValueError, OSError):
            pass
    server = LocalServer(("127.0.0.1", args.port), Handler)
    url = f"http://127.0.0.1:{server.server_port}"
    print(f"Ozon SKU Lab: {url}", flush=True)
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
