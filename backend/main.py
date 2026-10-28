import asyncio
import sys
from pathlib import Path

# Add backend directory to Python sys.path
backend_dir = Path(__file__).resolve().parent
if str(backend_dir) not in sys.path:
    sys.path.insert(0, str(backend_dir))

import httpx
from firebase_functions import https_fn, tasks_fn
from firebase_functions.options import RateLimits, RetryConfig
from werkzeug.wrappers import Response
from app import app, run_fetch_job_chunk

# In-process ASGI transport connecting directly to the FastAPI app
_transport = httpx.ASGITransport(app=app)


async def _dispatch_asgi(
    method: str,
    path: str,
    query_string: str,
    headers: dict,
    body: bytes,
) -> httpx.Response:
    url = path
    if query_string:
        url = f"{path}?{query_string}"
    async with httpx.AsyncClient(transport=_transport, base_url="http://app.internal") as client:
        return await client.request(
            method=method,
            url=url,
            headers=headers,
            content=body,
            timeout=535.0,
        )


@https_fn.on_request(memory=512, timeout_sec=540)
def api(req: https_fn.Request) -> https_fn.Response:
    """Firebase Cloud Functions v2 HTTP handler for FastAPI application."""
    path = req.path or "/"

    # Normalize double /api/api/... to /api/... if client sent redundant prefix
    while path.startswith("/api/api"):
        path = path.replace("/api/api", "/api", 1)

    # Ensure path starts with /api for FastAPI route matching
    if not path.startswith("/api"):
        path = f"/api{path}" if path.startswith("/") else f"/api/{path}"

    headers = dict(req.headers)
    headers.pop("Host", None)
    headers.pop("host", None)

    try:
        resp = asyncio.run(
            _dispatch_asgi(
                method=req.method,
                path=path,
                query_string=req.query_string.decode("utf-8") if req.query_string else "",
                headers=headers,
                body=req.get_data(),
            )
        )

        # Filter hop-by-hop headers and deduplicate
        out_headers = {}
        for k, v in resp.headers.items():
            if k.lower() in ("content-length", "content-encoding", "transfer-encoding", "connection"):
                continue
            out_headers[k] = v

        return Response(
            response=resp.content,
            status=resp.status_code,
            headers=out_headers,
        )
    except Exception as e:
        origin = req.headers.get("Origin", "*")
        return Response(
            response=f'{{"detail": "Internal server error: {str(e)}"}}',
            status=500,
            content_type="application/json",
            headers={
                "Access-Control-Allow-Origin": origin,
                "Access-Control-Allow-Credentials": "true",
                "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS, HEAD",
                "Access-Control-Allow-Headers": "*",
            },
        )


@tasks_fn.on_task_dispatched(
    # A busy or crashed chunk raises, so Cloud Tasks retries it after the lease expires.
    retry_config=RetryConfig(max_attempts=5, min_backoff_seconds=120),
    rate_limits=RateLimits(max_concurrent_dispatches=20),
    memory=1024,
    timeout_sec=1800,
)
def fetchjob(req: tasks_fn.CallableRequest) -> None:
    """Runs one chunk of a background fetch job (see run_fetch_job_chunk in app.py)."""
    job_id = (req.data or {}).get("job_id")
    if job_id:
        asyncio.run(run_fetch_job_chunk(job_id))
