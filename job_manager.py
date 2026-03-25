"""In-memory job queue with thread-safe state management."""

import threading
import uuid
from datetime import datetime
from typing import Optional


class JobManager:
    """Manages conversion jobs with thread-safe status tracking."""

    STATUSES = ('pending', 'converting', 'converted', 'verifying', 'done', 'error')

    def __init__(self):
        self._jobs = {}
        self._lock = threading.Lock()

    def create_job(self, pdf_name: str, pdf_path: str) -> dict:
        job_id = str(uuid.uuid4())[:8]
        job = {
            'id': job_id,
            'pdf_name': pdf_name,
            'pdf_path': pdf_path,
            'template_name': None,
            'template_path': None,
            'status': 'pending',
            'output_path': None,
            'error': None,
            'ai_output': None,
            'needs_verification': None,
            'force_verify': False,
            'created_at': datetime.now().isoformat(),
            'updated_at': datetime.now().isoformat(),
        }
        with self._lock:
            self._jobs[job_id] = job
        return job

    def update_job(self, job_id: str, **kwargs) -> Optional[dict]:
        with self._lock:
            job = self._jobs.get(job_id)
            if not job:
                return None
            for k, v in kwargs.items():
                if k in job:
                    job[k] = v
            job['updated_at'] = datetime.now().isoformat()
            return dict(job)

    def get_job(self, job_id: str) -> Optional[dict]:
        with self._lock:
            job = self._jobs.get(job_id)
            return dict(job) if job else None

    def get_all_jobs(self) -> list:
        with self._lock:
            return [dict(j) for j in self._jobs.values()]

    def get_pending_jobs(self) -> list:
        with self._lock:
            return [dict(j) for j in self._jobs.values() if j['status'] == 'pending']

    def delete_job(self, job_id: str) -> bool:
        with self._lock:
            return self._jobs.pop(job_id, None) is not None
