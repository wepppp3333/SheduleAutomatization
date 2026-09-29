import hmac
import os
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from uuid import uuid4

from fastapi import FastAPI, Header, HTTPException


BASE_DIR = Path(__file__).resolve().parent
AUTOMATION_SCRIPT = BASE_DIR / "barco_open_chrome.py"
API_TOKEN = os.getenv("BARCO_API_TOKEN", "")

app = FastAPI(title="Barco Schedule Automation", docs_url=None, redoc_url=None)

state_lock = threading.Lock()
state = {
    "job_id": None,
    "status": "idle",
    "started_at": None,
    "finished_at": None,
    "exit_code": None,
    "error": None,
}
process = None


def require_api_token(x_api_key: str):
    if not API_TOKEN:
        raise HTTPException(
            status_code=503,
            detail="BARCO_API_TOKEN is not configured on this computer",
        )
    if not hmac.compare_digest(x_api_key, API_TOKEN):
        raise HTTPException(status_code=401, detail="Invalid API token")


def run_automation(job_id: str):
    global process

    try:
        process = subprocess.Popen(
            [sys.executable, str(AUTOMATION_SCRIPT)],
            cwd=BASE_DIR,
        )
        exit_code = process.wait()
        with state_lock:
            state["status"] = "success" if exit_code == 0 else "failed"
            state["exit_code"] = exit_code
    except Exception as error:
        with state_lock:
            state["status"] = "failed"
            state["error"] = str(error)
    finally:
        with state_lock:
            state["finished_at"] = datetime.now().isoformat(timespec="seconds")
        process = None


@app.get("/health")
def health():
    return {
        "status": "ok",
        "computer": os.environ.get("COMPUTERNAME", "unknown"),
    }


@app.get("/status")
def get_status(x_api_key: str = Header(default="")):
    require_api_token(x_api_key)
    with state_lock:
        return dict(state)


@app.post("/run-schedule", status_code=202)
def run_schedule(x_api_key: str = Header(default="")):
    require_api_token(x_api_key)

    if not AUTOMATION_SCRIPT.exists():
        raise HTTPException(status_code=500, detail="Automation script not found")

    with state_lock:
        if state["status"] == "running":
            raise HTTPException(
                status_code=409,
                detail={"message": "Automation is already running", "job_id": state["job_id"]},
            )

        job_id = str(uuid4())
        state.update(
            {
                "job_id": job_id,
                "status": "running",
                "started_at": datetime.now().isoformat(timespec="seconds"),
                "finished_at": None,
                "exit_code": None,
                "error": None,
            }
        )

    worker = threading.Thread(target=run_automation, args=(job_id,), daemon=True)
    worker.start()
    return {"status": "started", "job_id": job_id}
