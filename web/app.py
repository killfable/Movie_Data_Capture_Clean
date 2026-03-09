import asyncio
import os
import shlex
import signal
import subprocess
import sys
import threading
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field


REPO_ROOT = Path(__file__).resolve().parent.parent
MAIN_PY = REPO_ROOT / "main.py"
LOG_DIR = REPO_ROOT / "logs" / "web_tasks"
LOG_DIR.mkdir(parents=True, exist_ok=True)


@dataclass
class TaskInfo:
    task_id: str
    command: List[str]
    created_at: str
    status: str = "pending"
    pid: Optional[int] = None
    return_code: Optional[int] = None
    finished_at: Optional[str] = None
    log_file: str = ""
    error: Optional[str] = None
    process: Optional[subprocess.Popen] = field(default=None, repr=False)


class TaskOptionsRequest(BaseModel):
    conf_file_path: str = ""
    over_config: List[str] = Field(default_factory=list)


class SearchRequest(TaskOptionsRequest):
    number: str = Field(min_length=1)


class SpecifyFileRequest(TaskOptionsRequest):
    file_path: str = Field(min_length=1)


class ScrapingUrlRequest(TaskOptionsRequest):
    url: str = Field(min_length=1)
    xlsx_file: str = Field(min_length=1)


class StartTaskResponse(BaseModel):
    task_id: str
    status: str
    command: str


class TaskManager:
    def __init__(self) -> None:
        self.tasks: Dict[str, TaskInfo] = {}
        self._lock = threading.Lock()

    def create_task(self, mode_args: List[str], options: Optional[TaskOptionsRequest] = None) -> TaskInfo:
        task_id = uuid.uuid4().hex[:12]
        log_file = str(LOG_DIR / f"{task_id}.log")
        cmd = [sys.executable, str(MAIN_PY)]

        if options:
            if options.conf_file_path:
                conf_path = Path(options.conf_file_path).expanduser()
                if not conf_path.exists():
                    raise ValueError(f"Config file not found: {options.conf_file_path}")
                cmd.extend(["--conf", str(conf_path)])
            for conf in options.over_config:
                cmd.extend(["--over-config", conf])

        cmd.extend(mode_args)

        task = TaskInfo(
            task_id=task_id,
            command=cmd,
            created_at=datetime.utcnow().isoformat() + "Z",
            status="running",
            log_file=log_file,
        )
        with self._lock:
            self.tasks[task_id] = task

        try:
            self._start_process(task)
        except Exception as exc:
            with self._lock:
                task.status = "failed"
                task.error = str(exc)
                task.finished_at = datetime.utcnow().isoformat() + "Z"
            raise
        return task

    def _start_process(self, task: TaskInfo) -> None:
        log_fp = open(task.log_file, "w", encoding="utf-8")
        process = subprocess.Popen(
            task.command,
            cwd=str(REPO_ROOT),
            stdout=log_fp,
            stderr=subprocess.STDOUT,
            text=True,
        )
        task.process = process
        task.pid = process.pid

        def wait_for_exit() -> None:
            return_code = process.wait()
            log_fp.close()
            with self._lock:
                task.return_code = return_code
                task.finished_at = datetime.utcnow().isoformat() + "Z"
                if task.status != "stopped":
                    task.status = "success" if return_code == 0 else "failed"

        threading.Thread(target=wait_for_exit, daemon=True).start()

    def stop_task(self, task_id: str) -> TaskInfo:
        with self._lock:
            task = self.tasks.get(task_id)
            if task is None:
                raise KeyError(task_id)
            if task.process is None or task.process.poll() is not None:
                return task

            try:
                task.process.send_signal(signal.SIGINT)
                task.status = "stopped"
            except Exception as exc:
                task.error = str(exc)
                raise
        return task

    def get_task(self, task_id: str) -> TaskInfo:
        with self._lock:
            task = self.tasks.get(task_id)
        if task is None:
            raise KeyError(task_id)
        return task

    def list_tasks(self) -> List[TaskInfo]:
        with self._lock:
            tasks = list(self.tasks.values())
        return sorted(tasks, key=lambda x: x.created_at, reverse=True)


app = FastAPI(title="Movie Data Capture Web")
app.mount("/static", StaticFiles(directory=str(REPO_ROOT / "web" / "static")), name="static")
manager = TaskManager()


def to_public(task: TaskInfo) -> Dict[str, Optional[str]]:
    return {
        "task_id": task.task_id,
        "status": task.status,
        "created_at": task.created_at,
        "finished_at": task.finished_at,
        "pid": task.pid,
        "return_code": task.return_code,
        "command": " ".join(shlex.quote(part) for part in task.command),
        "log_file": task.log_file,
        "error": task.error,
    }


def start_task(mode_args: List[str], options: Optional[TaskOptionsRequest] = None) -> StartTaskResponse:
    try:
        task = manager.create_task(mode_args, options)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))

    return StartTaskResponse(
        task_id=task.task_id,
        status=task.status,
        command=" ".join(shlex.quote(part) for part in task.command),
    )


@app.get("/")
async def index() -> FileResponse:
    return FileResponse(str(REPO_ROOT / "web" / "static" / "index.html"))


@app.get("/api/health")
async def health() -> Dict[str, str]:
    return {"status": "ok"}


@app.get("/api/tasks")
async def list_tasks() -> List[Dict[str, Optional[str]]]:
    return [to_public(t) for t in manager.list_tasks()]


@app.get("/api/tasks/{task_id}")
async def get_task(task_id: str) -> Dict[str, Optional[str]]:
    try:
        return to_public(manager.get_task(task_id))
    except KeyError:
        raise HTTPException(status_code=404, detail="Task not found")


@app.get("/api/tasks/{task_id}/log")
async def get_log(task_id: str) -> Dict[str, str]:
    try:
        task = manager.get_task(task_id)
    except KeyError:
        raise HTTPException(status_code=404, detail="Task not found")

    if not os.path.exists(task.log_file):
        return {"task_id": task_id, "log": ""}

    def read_file() -> str:
        with open(task.log_file, "r", encoding="utf-8", errors="replace") as fp:
            return fp.read()

    content = await asyncio.to_thread(read_file)
    return {"task_id": task_id, "log": content}


@app.post("/api/tasks/{task_id}/stop")
async def stop_task(task_id: str) -> Dict[str, Optional[str]]:
    try:
        task = manager.stop_task(task_id)
        return to_public(task)
    except KeyError:
        raise HTTPException(status_code=404, detail="Task not found")


@app.post("/api/tasks/search", response_model=StartTaskResponse)
async def run_search(req: SearchRequest) -> StartTaskResponse:
    return start_task(["--search", req.number], req)


@app.post("/api/tasks/list-movie", response_model=StartTaskResponse)
async def run_list_movie(req: TaskOptionsRequest) -> StartTaskResponse:
    return start_task(["--list-movie"], req)


@app.post("/api/tasks/specify-file", response_model=StartTaskResponse)
async def run_specify_file(req: SpecifyFileRequest) -> StartTaskResponse:
    return start_task(["--specify-file", req.file_path], req)


@app.post("/api/tasks/scraping-url", response_model=StartTaskResponse)
async def run_scraping_url(req: ScrapingUrlRequest) -> StartTaskResponse:
    return start_task(["--scraping-url", req.url, req.xlsx_file], req)


@app.post("/api/tasks/rate", response_model=StartTaskResponse)
async def run_rate(req: TaskOptionsRequest) -> StartTaskResponse:
    return start_task(["--rate"], req)
