import asyncio
import copy
import csv
import difflib
import functools
import io
import json
import os
import re
import smtplib
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Awaitable, Callable, Dict, List, Literal, Optional, Tuple

import firebase_admin
import gspread
import httpx
import openpyxl
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from firebase_admin import auth as firebase_auth
from firebase_admin import credentials, firestore
from firebase_admin import functions as admin_functions
from google import genai
from pydantic import BaseModel, Field

# Load environment variables from backend/.env (and .env.example as fallback)
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env", override=True)
load_dotenv(BASE_DIR / ".env.example", override=False)

# Initialize Firebase Admin SDK
SERVICE_ACCOUNT_PATH = BASE_DIR / "serviceAccountKey.json"
if not firebase_admin._apps:
    if SERVICE_ACCOUNT_PATH.exists():
        cred = credentials.Certificate(str(SERVICE_ACCOUNT_PATH))
        firebase_admin.initialize_app(cred)
    else:
        firebase_admin.initialize_app()

db = firestore.client()

app = FastAPI(
    title="Data Mining Agentic AI Backend",
    description="FastAPI backend powered by Gemini 3.8 Flash & Antigravity Agent with Firebase Auth, Cloud Firestore, and Two-Way Google Sheets Sync.",
    version="1.1.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------------
# Pydantic Schemas
# ---------------------------------------------------------------------------

class ColumnTemplate(BaseModel):
    id: str = Field(default_factory=lambda: f"col-{uuid.uuid4().hex[:8]}")
    name: str
    special_instructions: str = ""
    sort: Optional[Literal["Ascending", "Descending"]] = "Ascending"
    query_type: Literal["Main", "Attribute"] = "Attribute"
    data_type: Literal["String", "Number", "Date"] = "String"
    filter_op: Literal["Less Than", "Greater Than", "Equals", "Contains", "None"] = "Less Than"
    filter_value: str = ""
    put_link: bool = False
    retrieve_all: bool = False
    retrieve_all_mode: Literal["bullet_points", "more_rows"] = "bullet_points"


class SheetTab(BaseModel):
    id: str
    name: str
    gid: Optional[str] = None
    headers: List[str] = Field(default_factory=list)
    rows: List[Dict[str, Any]] = Field(default_factory=list)


class MCPConnector(BaseModel):
    id: str
    name: str
    icon: str  # "linkedin", "handshake", "custom"
    enabled: bool = True
    server_url: Optional[str] = None
    description: Optional[str] = None


class ProjectCreateRequest(BaseModel):
    name: str
    description: Optional[str] = ""


class ProjectUpdateRequest(BaseModel):
    name: Optional[str] = None
    existing_data_source: Optional[str] = None
    existing_data_type: Optional[Literal["csv", "gsheets", "none"]] = None
    active_tab_id: Optional[str] = None
    tabs: Optional[List[SheetTab]] = None
    columns: Optional[List[ColumnTemplate]] = None
    mcps: Optional[List[MCPConnector]] = None
    new_rows: Optional[List[Dict[str, Any]]] = None
    sync_to_gsheet: bool = False
    sync_to_gsheet: bool = False


class GSheetImportRequest(BaseModel):
    url: str
    tab_id: Optional[str] = None
    sync_columns: bool = True


class GSheetSyncRequest(BaseModel):
    direction: Literal["pull", "push"] = "pull"
    tab_id: Optional[str] = None


class CellUpdateRequest(BaseModel):
    tab_id: str
    row_index: int
    column_name: str
    value: Any
    sync_to_gsheet: bool = True


class AddRowRequest(BaseModel):
    tab_id: str
    row: Optional[Dict[str, Any]] = None
    sync_to_gsheet: bool = True


class FetchDataRequest(BaseModel):
    query_prompt: Optional[str] = ""
    tab_id: Optional[str] = None
    row_count: int = Field(default=10, ge=1, le=1000)
    append_rows: bool = True
    agent_mode: Literal["gemini-3.8-flash", "antigravity-agent", "gemini-4-argon"] = "gemini-3.8-flash"
    gemini_api_key: Optional[str] = None
    # Client-generated id so progress polling can ignore state left over from a previous run.
    run_id: Optional[str] = None


# ---------------------------------------------------------------------------
# Firebase Authentication Dependency
# ---------------------------------------------------------------------------

async def get_current_user(authorization: Optional[str] = Header(default=None)) -> Dict[str, Any]:
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(
            status_code=401,
            detail="Missing or invalid Authorization header. Expected 'Bearer <Firebase ID Token>'.",
        )
    token = authorization.split("Bearer ", 1)[1].strip()
    try:
        decoded_token = firebase_auth.verify_id_token(token)
        return {
            "uid": decoded_token.get("uid"),
            "email": decoded_token.get("email", ""),
            "name": decoded_token.get("name", ""),
        }
    except Exception as exc:
        raise HTTPException(
            status_code=401,
            detail=f"Invalid or expired Firebase authentication token: {str(exc)}",
        )


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _default_mcps() -> List[Dict[str, Any]]:
    return [
        {
            "id": "mcp-linkedin",
            "name": "LinkedIn",
            "icon": "linkedin",
            "enabled": True,
            "server_url": "",
            "description": "Scrape professional profiles, roles, companies, and career histories.",
        },
        {
            "id": "mcp-handshake",
            "name": "Handshake",
            "icon": "handshake",
            "enabled": True,
            "server_url": "",
            "description": "Extract university talent directories, career fairs, and school outreach data.",
        },
    ]


# ---------------------------------------------------------------------------
# Google Sheets Two-Way Sync Engine (Multi-Tab Read & Write)
# ---------------------------------------------------------------------------

def extract_spreadsheet_id(url_or_id: str) -> Optional[str]:
    text = (url_or_id or "").strip()
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", text)
    if m:
        return m.group(1)
    if re.match(r"^[a-zA-Z0-9-_]{25,}$", text):
        return text
    return None


def _get_gspread_client() -> gspread.Client:
    return gspread.service_account(filename=str(SERVICE_ACCOUNT_PATH))


def _parse_2d_sheet_values(
    ws_title: str, ws_gid: Optional[str], raw_values: List[List[Any]], idx: int
) -> Tuple[Dict[str, Any], Dict[str, str], bool]:
    """
    Parses a 2D list of cell values from a worksheet into:
      - A SheetTab dict: {id, name, gid, headers, rows}
      - A rules_map: {COLUMN_NAME_UPPER: special_instruction_string} if this tab is a RULES! tab
      - is_rules_tab: bool
    """
    # Filter out completely empty rows
    non_empty_rows: List[List[str]] = []
    max_col_used = 0
    for r in raw_values:
        str_row = [
            (str(int(c)) if isinstance(c, float) and c.is_integer() else str(c).strip())
            if c is not None
            else ""
            for c in r
        ]
        if any(cell != "" for cell in str_row):
            non_empty_rows.append(str_row)
            for c_idx, cell in enumerate(str_row):
                if cell != "" and c_idx + 1 > max_col_used:
                    max_col_used = c_idx + 1

    tab_id = f"tab-gs-{ws_gid}" if ws_gid is not None else f"tab-{idx + 1}"

    if not non_empty_rows or max_col_used == 0:
        return (
            {
                "id": tab_id,
                "name": ws_title,
                "gid": str(ws_gid) if ws_gid is not None else None,
                "headers": [],
                "rows": [],
            },
            {},
            False,
        )

    trimmed_rows = [r[:max_col_used] + [""] * max(0, max_col_used - len(r)) for r in non_empty_rows]
    first_row = trimmed_rows[0]

    # Detect if this is a RULES / Instructions tab (e.g. "RULES!" where first two header cells are blank)
    is_rules_tab = ws_title.strip().upper().startswith("RULE") or (
        len(first_row) >= 2 and first_row[0] == "" and first_row[1] == ""
    )

    rules_map: Dict[str, str] = {}
    headers: List[str] = []

    if is_rules_tab:
        for c_i in range(max_col_used):
            h_val = first_row[c_i].strip()
            if not h_val:
                if c_i == 0:
                    h_val = "Field / Column"
                elif c_i == 1:
                    h_val = "Rule / Instructions"
                else:
                    h_val = f"Column {c_i + 1}"
            headers.append(h_val)

        parsed_rows: List[Dict[str, Any]] = []
        for r in trimmed_rows[1:]:
            row_dict = {headers[c_i]: r[c_i] for c_i in range(len(headers))}
            parsed_rows.append(row_dict)
            if len(r) >= 2 and r[0].strip() and r[1].strip():
                rules_map[r[0].strip().upper()] = r[1].strip()

        return (
            {
                "id": tab_id,
                "name": ws_title,
                "gid": str(ws_gid) if ws_gid is not None else None,
                "headers": headers,
                "rows": parsed_rows,
            },
            rules_map,
            True,
        )

    # Standard Data Worksheet (e.g., Undergraduate, Graduate, Non-Eng/Tech)
    # Trim trailing columns that have no header name in the first row
    header_col_count = max_col_used
    while header_col_count > 1 and first_row[header_col_count - 1].strip() == "":
        header_col_count -= 1

    seen_headers: Dict[str, int] = {}
    for c_i in range(header_col_count):
        raw_h = first_row[c_i].strip() or f"Column {c_i + 1}"
        if raw_h in seen_headers:
            seen_headers[raw_h] += 1
            raw_h = f"{raw_h} ({seen_headers[raw_h]})"
        else:
            seen_headers[raw_h] = 1
        headers.append(raw_h)

    parsed_rows = []
    for r in trimmed_rows[1:]:
        row_dict = {headers[c_i]: r[c_i] for c_i in range(len(headers))}
        if any(str(v).strip() != "" for v in row_dict.values()):
            parsed_rows.append(row_dict)

    return (
        {
            "id": tab_id,
            "name": ws_title,
            "gid": str(ws_gid) if ws_gid is not None else None,
            "headers": headers,
            "rows": parsed_rows,
        },
        {},
        False,
    )


async def pull_all_tabs_from_google_sheet(
    sheet_url: str, existing_columns: Optional[List[Dict[str, Any]]] = None
) -> Dict[str, Any]:
    """
    Pulls ALL tabs (worksheets), headers, rows, and column rules from a Google Sheet.
    Uses `gspread` with serviceAccountKey.json first, falling back to `.xlsx` export.
    """
    sheet_id = extract_spreadsheet_id(sheet_url)
    if not sheet_id:
        raise HTTPException(
            status_code=400,
            detail="Invalid Google Sheets URL. Expected format: https://docs.google.com/spreadsheets/d/<ID>/edit",
        )

    gid_match = re.search(r"[#&?]gid=([0-9]+)", sheet_url)
    requested_gid = gid_match.group(1) if gid_match else None

    tabs: List[Dict[str, Any]] = []
    combined_rules: Dict[str, str] = {}
    primary_data_headers: List[str] = []
    spreadsheet_title = ""
    can_write = False

    # 1. Try gspread with service account (supports full read/write & exact gids)
    try:
        gc = _get_gspread_client()
        sh = gc.open_by_key(sheet_id)
        spreadsheet_title = sh.title
        worksheets = sh.worksheets()
        for idx, ws in enumerate(worksheets):
            raw_vals = ws.get_all_values()
            tab_obj, rules_map, is_rules = _parse_2d_sheet_values(ws.title, str(ws.id), raw_vals, idx)
            tabs.append(tab_obj)
            combined_rules.update(rules_map)
            if not is_rules and tab_obj["headers"] and not primary_data_headers:
                primary_data_headers = tab_obj["headers"]
        can_write = True
    except Exception:
        # 2. Fallback to public XLSX export (gets all tabs & rows even without service account access)
        xlsx_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
        async with httpx.AsyncClient(timeout=25.0, follow_redirects=True) as client:
            resp = await client.get(xlsx_url)
            if resp.status_code != 200:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"Could not access Google Sheet ({resp.status_code}). "
                        "Make sure the link is shared with 'Anyone with the link' or with "
                        "firebase-adminsdk-fbsvc@dr-system-c7488.iam.gserviceaccount.com."
                    ),
                )
            wb = openpyxl.load_workbook(io.BytesIO(resp.content), data_only=True)
            for idx, name in enumerate(wb.sheetnames):
                ws = wb[name]
                raw_vals = [list(r) for r in ws.iter_rows(values_only=True)]
                tab_obj, rules_map, is_rules = _parse_2d_sheet_values(name, str(idx), raw_vals, idx)
                tabs.append(tab_obj)
                combined_rules.update(rules_map)
                if not is_rules and tab_obj["headers"] and not primary_data_headers:
                    primary_data_headers = tab_obj["headers"]

    if not tabs:
        raise HTTPException(status_code=400, detail="No worksheets found in the Google Sheet.")

    if not primary_data_headers:
        primary_data_headers = tabs[0].get("headers", [])

    # Build synced Data Template columns from primary_data_headers + RULES! instructions
    existing_col_by_name = {
        c.get("name", "").strip().upper(): c for c in (existing_columns or [])
    }
    synced_columns: List[Dict[str, Any]] = []
    for idx, h in enumerate(primary_data_headers):
        h_upper = h.strip().upper()
        prev = existing_col_by_name.get(h_upper, {})
        rule_instr = combined_rules.get(h_upper) or prev.get("special_instructions", "")

        # Infer sensible defaults
        dtype: Literal["String", "Number", "Date"] = prev.get("data_type", "String")
        if "DATE" in h_upper:
            dtype = "Date"

        rule_lower = (rule_instr or "").lower()
        inferred_put_link = bool(
            prev.get("put_link")
            if "put_link" in prev
            else ("link to their profile" in rule_lower or "cntrl+k" in rule_lower or "ctrl+k" in rule_lower)
        )

        synced_columns.append({
            "id": prev.get("id") or f"col-{uuid.uuid4().hex[:8]}",
            "name": h,
            "special_instructions": rule_instr,
            "sort": prev.get("sort") if "sort" in prev else ("Ascending" if idx == 0 else None),
            "query_type": "Main" if idx == 0 else prev.get("query_type", "Attribute"),
            "data_type": dtype,
            "filter_op": prev.get("filter_op", "None"),
            "filter_value": prev.get("filter_value", ""),
            "put_link": inferred_put_link,
            "retrieve_all": bool(prev.get("retrieve_all", False)),
            "retrieve_all_mode": prev.get("retrieve_all_mode", "bullet_points"),
        })

    # Choose best active_tab_id: match requested_gid if present, else first non-RULES tab, else first tab
    active_tab_id = tabs[0]["id"]
    if requested_gid:
        for t in tabs:
            if t.get("gid") == requested_gid:
                active_tab_id = t["id"]
                break
    else:
        for t in tabs:
            if not t["name"].strip().upper().startswith("RULE"):
                active_tab_id = t["id"]
                break

    return {
        "spreadsheet_title": spreadsheet_title,
        "tabs": tabs,
        "columns": synced_columns,
        "active_tab_id": active_tab_id,
        "can_write": can_write,
    }


def push_tab_to_google_sheet(
    sheet_url: str,
    tab: Dict[str, Any],
    fallback_columns: List[Dict[str, Any]],
    append_only_rows: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """
    Writes a tab (or appends new rows) directly back to the live Google Sheet via gspread.
    Creates the worksheet tab in Google Sheets if it doesn't exist yet.
    """
    sheet_id = extract_spreadsheet_id(sheet_url)
    if not sheet_id:
        return {"synced": False, "reason": "No valid Google Sheet ID"}

    try:
        gc = _get_gspread_client()
        sh = gc.open_by_key(sheet_id)
        worksheets = sh.worksheets()

        target_ws = None
        tab_gid = str(tab.get("gid") or "")
        tab_name = (tab.get("name") or "Sheet1").strip()

        for ws in worksheets:
            if tab_gid and str(ws.id) == tab_gid:
                target_ws = ws
                break
        if target_ws is None:
            for ws in worksheets:
                if ws.title.strip().lower() == tab_name.lower():
                    target_ws = ws
                    break

        headers: List[str] = tab.get("headers") or [c.get("name", "") for c in fallback_columns if c.get("name")]
        rows: List[Dict[str, Any]] = tab.get("rows", [])
        if not headers and rows:
            headers = list(rows[0].keys())

        # Create worksheet in Google Sheet if user added a new tab in the web app
        if target_ws is None:
            target_ws = sh.add_worksheet(
                title=tab_name,
                rows=max(100, len(rows) + 20),
                cols=max(10, len(headers) + 2),
            )
            tab["gid"] = str(target_ws.id)
            if headers:
                target_ws.update(values=[headers], range_name="A1", value_input_option="USER_ENTERED")
        else:
            # Sync tab rename if changed
            if target_ws.title != tab_name:
                try:
                    target_ws.update_title(tab_name)
                except Exception:
                    pass

        if append_only_rows is not None and len(append_only_rows) > 0:
            # Non-destructive append of newly scraped/added rows
            existing_vals = target_ws.get_all_values()
            if not existing_vals and headers:
                target_ws.append_row(headers, value_input_option="USER_ENTERED")
            elif existing_vals:
                sheet_headers = [h.strip() for h in existing_vals[0] if h.strip()]
                if sheet_headers:
                    headers = sheet_headers

            rows_to_append = [
                [str(r.get(h, "")) if r.get(h) is not None else "" for h in headers]
                for r in append_only_rows
            ]
            target_ws.append_rows(rows_to_append, value_input_option="USER_ENTERED")
            return {
                "synced": True,
                "mode": "append",
                "worksheet": target_ws.title,
                "rows_written": len(rows_to_append),
            }

        # Full worksheet update
        matrix: List[List[Any]] = []
        if headers:
            matrix.append(headers)
        for r in rows:
            matrix.append([str(r.get(h, "")) if r.get(h) is not None else "" for h in headers])

        target_ws.clear()
        if matrix:
            target_ws.update(values=matrix, range_name="A1", value_input_option="USER_ENTERED")

        return {
            "synced": True,
            "mode": "full_update",
            "worksheet": target_ws.title,
            "rows_written": len(rows),
        }
    except Exception as exc:
        return {"synced": False, "reason": str(exc)}


def update_single_cell_in_google_sheet(
    sheet_url: str,
    tab: Dict[str, Any],
    row_index: int,
    column_name: str,
    value: Any,
    fallback_columns: List[Dict[str, Any]],
) -> Dict[str, Any]:
    """
    Updates a single cell directly in the live Google Sheet without touching any other cells/formatting.
    """
    sheet_id = extract_spreadsheet_id(sheet_url)
    if not sheet_id:
        return {"synced": False, "reason": "No valid Google Sheet ID"}

    try:
        gc = _get_gspread_client()
        sh = gc.open_by_key(sheet_id)
        worksheets = sh.worksheets()
        target_ws = None
        tab_gid = str(tab.get("gid") or "")
        tab_name = (tab.get("name") or "").strip()

        for ws in worksheets:
            if (tab_gid and str(ws.id) == tab_gid) or ws.title.strip().lower() == tab_name.lower():
                target_ws = ws
                break

        if target_ws is None:
            return push_tab_to_google_sheet(sheet_url, tab, fallback_columns)

        headers: List[str] = tab.get("headers") or [c.get("name", "") for c in fallback_columns if c.get("name")]
        if column_name not in headers:
            return push_tab_to_google_sheet(sheet_url, tab, fallback_columns)

        col_idx_1based = headers.index(column_name) + 1
        row_idx_1based = row_index + 2  # +1 for 1-based, +1 for header row
        target_ws.update_cell(row_idx_1based, col_idx_1based, str(value) if value is not None else "")
        return {"synced": True, "mode": "cell_update", "worksheet": target_ws.title}
    except Exception as exc:
        return {"synced": False, "reason": str(exc)}


# ---------------------------------------------------------------------------
# Default / Seed Projects
# ---------------------------------------------------------------------------

def _build_starter_projects(uid: str, email: str) -> List[Dict[str, Any]]:
    now = _now_iso()
    default_headers = ["Name", "Job", "Year", "Location"]
    return [
        {
            "name": "Employee-data",
            "owner_uid": uid,
            "owner_email": email,
            "created_at": now,
            "last_edited": now,
            "existing_data_source": "company_directory_seed.csv",
            "existing_data_type": "csv",
            "active_tab_id": "tab-1",
            "columns": [
                {
                    "id": "col-name",
                    "name": "Name",
                    "special_instructions": "Extract full name of the professional or lead.",
                    "sort": "Ascending",
                    "query_type": "Main",
                    "data_type": "String",
                    "filter_op": "Less Than",
                    "filter_value": "",
                },
                {
                    "id": "col-job",
                    "name": "Job",
                    "special_instructions": "Current job title and department.",
                    "sort": None,
                    "query_type": "Attribute",
                    "data_type": "String",
                    "filter_op": "None",
                    "filter_value": "",
                },
                {
                    "id": "col-year",
                    "name": "Year",
                    "special_instructions": "Year joined or graduation year.",
                    "sort": None,
                    "query_type": "Attribute",
                    "data_type": "Number",
                    "filter_op": "None",
                    "filter_value": "",
                },
                {
                    "id": "col-location",
                    "name": "Location",
                    "special_instructions": "City and State / Country.",
                    "sort": None,
                    "query_type": "Attribute",
                    "data_type": "String",
                    "filter_op": "None",
                    "filter_value": "",
                },
            ],
            "tabs": [
                {
                    "id": "tab-1",
                    "name": "Tab 1",
                    "headers": default_headers,
                    "rows": [
                        {"Name": "Elena Rostova", "Job": "Senior Data Engineer", "Year": 2022, "Location": "San Francisco, CA"},
                        {"Name": "Marcus Vance", "Job": "AI Product Manager", "Year": 2023, "Location": "New York, NY"},
                        {"Name": "Priya Patel", "Job": "Principal ML Scientist", "Year": 2021, "Location": "Seattle, WA"},
                        {"Name": "Sophia Chen", "Job": "Analytics Engineering Lead", "Year": 2024, "Location": "Austin, TX"},
                    ],
                },
                {
                    "id": "tab-2",
                    "name": "Tab 2",
                    "headers": default_headers,
                    "rows": [
                        {"Name": "David Kim", "Job": "VP of Engineering", "Year": 2020, "Location": "Boston, MA"},
                        {"Name": "Hannah Abbott", "Job": "Staff Data Analyst", "Year": 2023, "Location": "Chicago, IL"},
                    ],
                },
            ],
            "mcps": _default_mcps(),
            "last_agent_run": None,
        },
        {
            "name": "Finding-newsletters",
            "owner_uid": uid,
            "owner_email": email,
            "created_at": now,
            "last_edited": now,
            "existing_data_source": "tech_newsletters.csv",
            "existing_data_type": "csv",
            "active_tab_id": "tab-1",
            "columns": [
                {
                    "id": "col-nl-name",
                    "name": "Name",
                    "special_instructions": "Newsletter title and publisher name.",
                    "sort": "Ascending",
                    "query_type": "Main",
                    "data_type": "String",
                    "filter_op": "Less Than",
                    "filter_value": "",
                },
                {
                    "id": "col-nl-topic",
                    "name": "Job",
                    "special_instructions": "Target audience role or primary industry focus.",
                    "sort": None,
                    "query_type": "Attribute",
                    "data_type": "String",
                    "filter_op": "None",
                    "filter_value": "",
                },
                {
                    "id": "col-nl-year",
                    "name": "Year",
                    "special_instructions": "Year launched.",
                    "sort": None,
                    "query_type": "Attribute",
                    "data_type": "Number",
                    "filter_op": "None",
                    "filter_value": "",
                },
                {
                    "id": "col-nl-loc",
                    "name": "Location",
                    "special_instructions": "Publisher HQ or primary region.",
                    "sort": None,
                    "query_type": "Attribute",
                    "data_type": "String",
                    "filter_op": "None",
                    "filter_value": "",
                },
            ],
            "tabs": [
                {
                    "id": "tab-1",
                    "name": "Tab 1",
                    "headers": default_headers,
                    "rows": [
                        {"Name": "TLDR AI Newsletter", "Job": "AI & Software Engineers", "Year": 2023, "Location": "Seattle, WA"},
                        {"Name": "The Batch (DeepLearning.AI)", "Job": "ML Researchers & Practitioners", "Year": 2019, "Location": "Palo Alto, CA"},
                        {"Name": "Pragmatic Engineer", "Job": "Engineering Managers & Devs", "Year": 2021, "Location": "Amsterdam, NL"},
                    ],
                },
                {"id": "tab-2", "name": "Tab 2", "headers": default_headers, "rows": []},
            ],
            "mcps": _default_mcps(),
            "last_agent_run": None,
        },
        {
            "name": "School-outreach",
            "owner_uid": uid,
            "owner_email": email,
            "created_at": now,
            "last_edited": now,
            "existing_data_source": "university_career_centers.csv",
            "existing_data_type": "csv",
            "active_tab_id": "tab-1",
            "columns": [
                {
                    "id": "col-sc-name",
                    "name": "Name",
                    "special_instructions": "University name or Career Services director.",
                    "sort": "Ascending",
                    "query_type": "Main",
                    "data_type": "String",
                    "filter_op": "Less Than",
                    "filter_value": "",
                },
                {
                    "id": "col-sc-job",
                    "name": "Job",
                    "special_instructions": "Program or Career Center role.",
                    "sort": None,
                    "query_type": "Attribute",
                    "data_type": "String",
                    "filter_op": "None",
                    "filter_value": "",
                },
                {
                    "id": "col-sc-year",
                    "name": "Year",
                    "special_instructions": "Target cohort graduation year.",
                    "sort": None,
                    "query_type": "Attribute",
                    "data_type": "Number",
                    "filter_op": "None",
                    "filter_value": "",
                },
                {
                    "id": "col-sc-loc",
                    "name": "Location",
                    "special_instructions": "Campus city and state.",
                    "sort": None,
                    "query_type": "Attribute",
                    "data_type": "String",
                    "filter_op": "None",
                    "filter_value": "",
                },
            ],
            "tabs": [
                {
                    "id": "tab-1",
                    "name": "Tab 1",
                    "headers": default_headers,
                    "rows": [
                        {"Name": "UC Berkeley Career Engagement", "Job": "STEM & CS Recruiting", "Year": 2027, "Location": "Berkeley, CA"},
                        {"Name": "Georgia Tech Career Center", "Job": "Co-op & University Relations", "Year": 2027, "Location": "Atlanta, GA"},
                        {"Name": "UT Austin Texas Career Engagement", "Job": "Early Talent Outreach", "Year": 2026, "Location": "Austin, TX"},
                    ],
                },
                {"id": "tab-2", "name": "Tab 2", "headers": default_headers, "rows": []},
            ],
            "mcps": _default_mcps(),
            "last_agent_run": None,
        },
    ]


# ---------------------------------------------------------------------------
# Helper: Sort & Filter Rows According to Data Template
# ---------------------------------------------------------------------------

def _group_merged_rows(
    rows: List[Dict[str, Any]], columns: List[Dict[str, Any]]
) -> List[List[Dict[str, Any]]]:
    """
    Groups consecutive rows that belong to the same vertically-merged entity block
    (where row 0 has the primary column populated and continuation rows have the
    primary column blank `""` or share `_group_id`).
    """
    if not rows:
        return []
    main_col = columns[0].get("name", "") if columns else ""

    groups: List[List[Dict[str, Any]]] = []
    for r in rows:
        gid = r.get("_group_id")
        main_val = str(r.get(main_col, "") or "").strip() if main_col else ""
        has_other_vals = any(
            not str(k).startswith("_") and k != main_col and str(v or "").strip() != ""
            for k, v in r.items()
        )
        if groups and (
            (gid and gid == groups[-1][0].get("_group_id"))
            or (main_col and main_val == "" and has_other_vals)
        ):
            groups[-1].append(r)
        else:
            groups.append([r])
    return groups


def apply_template_transforms(
    rows: List[Dict[str, Any]], columns: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    if not rows or not columns:
        return rows

    groups = _group_merged_rows(list(rows), columns)

    for col in columns:
        col_name = col.get("name", "")
        op = col.get("filter_op", "None")
        val_str = str(col.get("filter_value", "")).strip()
        dtype = col.get("data_type", "String")

        if not col_name or not val_str or op == "None":
            continue

        def _row_matches(r: Dict[str, Any]) -> bool:
            cell = r.get(col_name, "")
            if cell is None:
                cell = ""
            cell_str = str(cell).strip()
            if not cell_str:
                return False
            try:
                if dtype == "Number":
                    c_num = float(re.sub(r"[^\d.-]", "", cell_str))
                    v_num = float(re.sub(r"[^\d.-]", "", val_str))
                    if op == "Less Than":
                        return c_num < v_num
                    if op == "Greater Than":
                        return c_num > v_num
                    if op == "Equals":
                        return c_num == v_num
                    if op == "Contains":
                        return val_str.lower() in cell_str.lower()
                else:
                    if op == "Less Than":
                        return cell_str.lower() < val_str.lower()
                    if op == "Greater Than":
                        return cell_str.lower() > val_str.lower()
                    if op == "Equals":
                        return cell_str.lower() == val_str.lower()
                    if op == "Contains":
                        return val_str.lower() in cell_str.lower()
            except Exception:
                return False
            return False

        groups = [g for g in groups if any(_row_matches(r) for r in g)]

    sort_cols = [c for c in columns if c.get("sort") in ("Ascending", "Descending")]
    sort_cols.sort(key=lambda c: 0 if c.get("query_type") == "Main" else 1)

    if sort_cols:
        primary_col = sort_cols[0]
        col_name = primary_col.get("name", "")
        reverse = primary_col.get("sort") == "Descending"
        dtype = primary_col.get("data_type", "String")

        def group_sort_key(group: List[Dict[str, Any]]):
            head = group[0] if group else {}
            val = head.get(col_name, "")
            if val is None or str(val).strip() == "":
                return (1, "")
            if dtype == "Number":
                try:
                    return (0, float(re.sub(r"[^\d.-]", "", str(val))))
                except Exception:
                    return (1, str(val).lower())
            return (0, str(val).lower())

        groups.sort(key=group_sort_key, reverse=reverse)

    flattened: List[Dict[str, Any]] = []
    for g in groups:
        flattened.extend(g)
    return flattened


# ---------------------------------------------------------------------------
# Helper: Direct Web Page Scraper for Explicit URLs
# ---------------------------------------------------------------------------

async def scrape_urls_context(urls: List[str]) -> str:
    snippets = []
    async with httpx.AsyncClient(timeout=12.0, follow_redirects=True) as http_client:
        for url in urls[:3]:
            if "docs.google.com/spreadsheets" in url:
                continue
            try:
                resp = await http_client.get(
                    url,
                    headers={
                        "User-Agent": (
                            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
                        )
                    },
                )
                if resp.status_code == 200:
                    soup = BeautifulSoup(resp.text, "html.parser")
                    for tag in soup(["script", "style", "noscript"]):
                        tag.decompose()
                    title = soup.title.string.strip() if soup.title and soup.title.string else url
                    text = " ".join(soup.get_text(separator=" ").split())[:4000]
                    snippets.append(f"URL: {url}\nTitle: {title}\nContent Excerpt: {text}")
            except Exception as exc:
                snippets.append(f"URL: {url} (Fetch notice: {exc})")
    return "\n\n---\n\n".join(snippets)


# ---------------------------------------------------------------------------
# Google Sheets Format Learner & Deduplication Engine
# ---------------------------------------------------------------------------

_WS_RE = re.compile(r"\s+")
_PAREN_RE = re.compile(r"\([^)]*\)")
_NON_ALNUM_RE = re.compile(r"[^a-z0-9\s]")


@functools.lru_cache(maxsize=65536)
def _normalize_entity_cached(s: str) -> str:
    return _WS_RE.sub(" ", s.strip().lower())


@functools.lru_cache(maxsize=65536)
def _normalize_name_core_cached(s: str) -> str:
    s = _PAREN_RE.sub(" ", s.strip().lower())
    s = _NON_ALNUM_RE.sub("", s)
    return " ".join(s.split())


def _normalize_entity_str(val: Any) -> str:
    if val is None:
        return ""
    return _normalize_entity_cached(str(val))


def _normalize_name_core(val: Any) -> str:
    if val is None:
        return ""
    return _normalize_name_core_cached(str(val))


# Fuzzy near-duplicate matching ("Jon Smith" vs "John Smith").
# Uses a deletion-neighborhood index (SymSpell-style): two strings within edit distance 1 always
# share at least one "delete one character" variant, so a lookup is ~len(name) dict probes instead
# of scanning every stored name. Hits are then verified with a similarity ratio so short names
# (where one edit is a big change) are not wrongly merged.
FUZZY_THRESHOLD = 0.92
FUZZY_MIN_LEN = 6
FUZZY_MAX_LEN = 40


def _delete_variants(core: str) -> set[str]:
    variants = {core}
    for i in range(len(core)):
        variants.add(core[:i] + core[i + 1:])
    return variants


class _FuzzyNameIndex:
    def __init__(self) -> None:
        self.index: Dict[str, List[str]] = {}

    def add(self, core: str) -> None:
        if not (FUZZY_MIN_LEN <= len(core) <= FUZZY_MAX_LEN):
            return
        for variant in _delete_variants(core):
            self.index.setdefault(variant, []).append(core)

    def has_similar(self, core: str) -> bool:
        if not (FUZZY_MIN_LEN <= len(core) <= FUZZY_MAX_LEN):
            return False
        checked: set[str] = set()
        for variant in _delete_variants(core):
            for other in self.index.get(variant, ()):
                if other in checked:
                    continue
                checked.add(other)
                sm = difflib.SequenceMatcher(None, core, other, autojunk=False)
                if sm.ratio() >= FUZZY_THRESHOLD:
                    return True
        return False


def _token_sorted_key(core: str) -> str:
    """Order-insensitive key so 'Smith John' matches 'John Smith'. Empty for single tokens."""
    tokens = core.split()
    if len(tokens) < 2:
        return ""
    return " ".join(sorted(tokens))


class DeduplicationEngine:
    """
    Guarantees that newly generated rows NEVER repeat any record from Existing Data
    (across all worksheet tabs in the project) and NEVER repeat within the new batch.

    Exact checks are O(1) hash-set lookups. Name-like columns additionally get an
    order-insensitive match and a blocked fuzzy match for near-duplicate spellings.
    """

    _ALIAS_KEYS = ("NAME", "FULL NAME", "EMAIL", "CONTACT EMAIL")
    _ALIAS_FUZZY_KEYS = ("NAME", "FULL NAME")

    def __init__(self, existing_rows: List[Dict[str, Any]], columns: List[Dict[str, Any]]):
        self.columns = columns
        self.col_names = [c.get("name", "") for c in columns if c.get("name")]
        # Primary unique-entity columns are identity columns (NAME, EMAIL, LINK/URL, or column 0).
        # Shared categorical columns like COMPANY, POSITION, JOB TYPE, MAJOR, YEAR, LOCATION can appear
        # across multiple different people, so they are checked via composite keys rather than solo keys.
        non_unique_shared_keywords = (
            "COMPANY", "EMPLOYER", "ORGANIZATION", "JOB", "POSITION", "ROLE", "TITLE",
            "MAJOR", "AREA OF STUDY", "DEGREE", "YEAR", "COHORT", "SALARY", "LOCATION", "CITY", "NOTES", "NOTE",
        )
        self.main_cols: List[str] = []
        for idx_c, c in enumerate(columns):
            cn = c.get("name", "")
            if not cn:
                continue
            cn_up = cn.strip().upper()
            is_shared_attr = any(k in cn_up for k in non_unique_shared_keywords)
            if any(k in cn_up for k in ("NAME", "EMAIL", "LINK", "URL", "NEWSLETTER", "PERSON", "CANDIDATE", "STUDENT")):
                if cn not in self.main_cols:
                    self.main_cols.append(cn)
            elif (idx_c == 0 or c.get("query_type") == "Main") and not is_shared_attr:
                if cn not in self.main_cols:
                    self.main_cols.append(cn)
        if not self.main_cols and self.col_names:
            self.main_cols = [self.col_names[0]]

        # Fuzzy matching only makes sense for human-name-like text, never emails/links/ids.
        self.fuzzy_cols: set[str] = {
            mc for mc in self.main_cols
            if not any(k in mc.strip().upper() for k in ("EMAIL", "LINK", "URL"))
        }

        self.seen_entities: set[str] = set()
        self.seen_name_cores: set[str] = set()
        self.seen_token_keys: set[str] = set()
        self.seen_row_sigs: set[Tuple[str, ...]] = set()
        self.seen_composite_keys: set[Tuple[str, ...]] = set()
        self._fuzzy = _FuzzyNameIndex()

        for r in existing_rows:
            self.register_existing(r)

    def _identities(
        self, row: Dict[str, Any], include_aliases: bool
    ) -> List[Tuple[str, str, bool]]:
        """Normalize each identity cell exactly once -> [(entity, name_core, fuzzy_eligible)]."""
        out: List[Tuple[str, str, bool]] = []
        seen: set[Tuple[str, str, bool]] = set()

        def add(raw: Any, fuzzy: bool) -> None:
            item = (_normalize_entity_str(raw), _normalize_name_core(raw), fuzzy)
            if item not in seen:
                seen.add(item)
                out.append(item)

        for mc in self.main_cols:
            add(row.get(mc, ""), mc in self.fuzzy_cols)
        if include_aliases:
            # Also catch rows whose headers differ slightly in casing/spacing
            for k, v in row.items():
                ku = k.strip().upper()
                if ku in self._ALIAS_KEYS:
                    add(v, ku in self._ALIAS_FUZZY_KEYS)
        return out

    def register_existing(self, row: Dict[str, Any]) -> None:
        if not isinstance(row, dict):
            return
        # Full row signature
        sig = tuple(_normalize_entity_str(row.get(c, "")) for c in self.col_names)
        if any(sig):
            self.seen_row_sigs.add(sig)

        # Primary entity values
        for ent, core, fuzzy in self._identities(row, include_aliases=True):
            if ent and len(ent) > 1:
                self.seen_entities.add(ent)
            if core and len(core) > 1:
                if core not in self.seen_name_cores:
                    self.seen_name_cores.add(core)
                    if fuzzy:
                        self._fuzzy.add(core)
                token_key = _token_sorted_key(core)
                if token_key:
                    self.seen_token_keys.add(token_key)

        # Composite signature across the first 4 columns
        comp = sig[:4]
        if any(comp):
            self.seen_composite_keys.add(comp)

    def is_duplicate(self, candidate: Dict[str, Any]) -> bool:
        if not isinstance(candidate, dict):
            return True
        sig = tuple(_normalize_entity_str(candidate.get(c, "")) for c in self.col_names)
        if not any(sig):
            return True
        if sig in self.seen_row_sigs:
            return True

        for ent, core, fuzzy in self._identities(candidate, include_aliases=False):
            if ent and ent in self.seen_entities:
                return True
            if core:
                if core in self.seen_name_cores:
                    return True
                token_key = _token_sorted_key(core)
                if token_key and token_key in self.seen_token_keys:
                    return True
                if fuzzy and self._fuzzy.has_similar(core):
                    return True

        comp = sig[:4]
        if any(comp) and comp in self.seen_composite_keys:
            return True

        return False

    def accept(self, candidate: Dict[str, Any]) -> bool:
        if self.is_duplicate(candidate):
            return False
        self.register_existing(candidate)
        return True


def learn_sheet_format_profile(
    target_tab: Dict[str, Any],
    all_tabs: List[Dict[str, Any]],
    columns: List[Dict[str, Any]],
) -> Dict[str, Dict[str, Any]]:
    """
    Inspects the existing data seen in the Google Sheet (or CSV) tab and learns
    per-column formatting patterns, dropdown/categorical values, date/number styles,
    and example values so newly produced rows match the exact sheet format.
    """
    rows = target_tab.get("rows", [])
    if not rows:
        for t in all_tabs:
            if not t.get("name", "").strip().upper().startswith("RULE") and t.get("rows"):
                rows = t.get("rows", [])
                break

    profiles: Dict[str, Dict[str, Any]] = {}
    for col in columns:
        c_name = col.get("name", "")
        c_dtype = col.get("data_type", "String")
        non_empty: List[str] = []
        for r in rows:
            val = r.get(c_name)
            if val is not None and str(val).strip() != "":
                non_empty.append(str(val).strip())

        total_rows = max(1, len(rows))
        fill_rate = len(non_empty) / total_rows

        freq: Dict[str, int] = {}
        for v in non_empty:
            # Skip instructional banner rows (very long explanatory text in row 0)
            if len(v) > 95 and "Ex." in v:
                continue
            freq[v] = freq.get(v, 0) + 1

        sorted_vals = sorted(freq.items(), key=lambda kv: kv[1], reverse=True)
        top_values = [k for k, _ in sorted_vals[:15]]
        distinct_count = len(freq)

        # Detect if categorical / dropdown (e.g., JOB TYPE, AREA OF STUDY, etc.)
        is_categorical = (
            len(non_empty) >= 5
            and distinct_count <= max(12, int(len(non_empty) * 0.15))
            and c_name.strip().upper() not in ("NAME", "FULL NAME", "EMAIL", "CONTACT EMAIL", "LINK", "URL")
        )

        # Detect date pattern
        date_pattern = "YYYY-MM-DD"
        if non_empty:
            sample0 = non_empty[0]
            if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", sample0):
                date_pattern = "MM/DD/YYYY"
            elif re.match(r"^\d{1,2}/\d{1,2}/\d{2}$", sample0):
                date_pattern = "MM/DD/YY"
            elif re.match(r"^[A-Za-z]+ \d{1,2}, \d{4}$", sample0):
                date_pattern = "Month D, YYYY"

        # Detect case style
        case_style = "normal"
        if len(non_empty) >= 3:
            if all(v.isupper() for v in non_empty[:20] if any(ch.isalpha() for ch in v)):
                case_style = "upper"
            elif all(v.islower() for v in non_empty[:20] if any(ch.isalpha() for ch in v)):
                case_style = "lower"

        profiles[c_name] = {
            "name": c_name,
            "data_type": c_dtype,
            "fill_rate": fill_rate,
            "is_categorical": is_categorical,
            "top_values": top_values,
            "sample_values": non_empty[:6],
            "date_pattern": date_pattern,
            "case_style": case_style,
        }
    return profiles


def sanitize_cell_for_gsheets(
    val: Any,
    col: Dict[str, Any],
    profile: Dict[str, Any],
    allow_multiline: bool = False,
) -> Any:
    """
    Formats a single cell value so it matches the Google Sheet's seen data style.
    When `allow_multiline` is True (e.g., for bullet-point lists in a single cell),
    internal newlines `\n` are preserved and quoted in TSV/HTML clipboard output.
    """
    if val is None:
        return ""
    dtype = col.get("data_type", "String")
    if dtype == "Number" and not allow_multiline and "•" not in str(val):
        if isinstance(val, (int, float)):
            return val
        cleaned = re.sub(r"[^\d.-]", "", str(val).strip())
        if cleaned:
            try:
                f_val = float(cleaned)
                return int(f_val) if f_val.is_integer() else f_val
            except Exception:
                pass
    raw_str = str(val).replace("\t", " ").replace("\r\n", "\n").replace("\r", "\n")
    if allow_multiline or (
        col.get("retrieve_all") and col.get("retrieve_all_mode", "bullet_points") == "bullet_points"
    ) or "\n•" in raw_str or raw_str.strip().startswith("•"):
        lines = [ln.strip() for ln in raw_str.split("\n") if ln.strip()]
        text = "\n".join(lines)
    else:
        text = raw_str.replace("\n", " ").strip()

    case_style = profile.get("case_style", "normal")
    if case_style == "upper":
        text = text.upper()
    elif case_style == "lower":
        text = text.lower()
    return text


def format_row_for_gsheets(
    raw_row: Dict[str, Any],
    columns: List[Dict[str, Any]],
    profiles: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]:
    """
    Orders keys in the exact Google Sheet column order, formats each cell value,
    and preserves `_links` and merged-row group metadata (`_group_id`, `_group_span`, `_group_index`).
    """
    formatted: Dict[str, Any] = {}
    upper_lookup = {str(k).strip().upper(): v for k, v in raw_row.items() if not str(k).startswith("_")}
    for col in columns:
        c_name = col.get("name", "")
        val = raw_row.get(c_name)
        if val is None:
            val = upper_lookup.get(c_name.strip().upper(), "")
        prof = profiles.get(c_name, {})
        allow_multi = bool(
            col.get("retrieve_all") and col.get("retrieve_all_mode", "bullet_points") == "bullet_points"
        ) or ("\n" in str(val or ""))
        formatted[c_name] = sanitize_cell_for_gsheets(val, col, prof, allow_multiline=allow_multi)

    # Preserve hyperlink metadata and merged-group metadata
    if isinstance(raw_row.get("_links"), dict) and raw_row["_links"]:
        formatted["_links"] = dict(raw_row["_links"])
    for meta_key in ("_group_id", "_group_span", "_group_index", "_source"):
        if meta_key in raw_row:
            formatted[meta_key] = raw_row[meta_key]
    return formatted


def rows_to_gsheets_tsv(
    rows: List[Dict[str, Any]],
    columns: List[Dict[str, Any]],
    include_headers: bool = False,
) -> str:
    """
    Converts rows into Tab-Separated Values (TSV) formatted for direct Cmd+V / Ctrl+V
    pasting into Google Sheets. Multi-line cells (such as bullet points) are wrapped
    in RFC-4180 double quotes so Google Sheets pastes all bullets inside a single cell.
    """
    headers = [c.get("name", "") for c in columns if c.get("name")]
    lines: List[str] = []
    if include_headers and headers:
        lines.append("\t".join(h.replace("\t", " ").replace("\n", " ") for h in headers))
    for r in rows:
        cells = []
        for h in headers:
            v = r.get(h, "")
            if v is None:
                v = ""
            s = str(v).replace("\t", " ").replace("\r\n", "\n").replace("\r", "\n").strip()
            if "\n" in s or '"' in s:
                escaped = s.replace('"', '""')
                cells.append(f'"{escaped}"')
            else:
                cells.append(s)
        lines.append("\t".join(cells))
    return "\n".join(lines)


def _row_matches_column_filters(row: Dict[str, Any], columns: List[Dict[str, Any]]) -> bool:
    for col in columns:
        col_name = col.get("name", "")
        op = col.get("filter_op", "None")
        val_str = str(col.get("filter_value", "")).strip()
        dtype = col.get("data_type", "String")

        if not col_name or not val_str or op == "None":
            continue

        cell = row.get(col_name, "")
        cell_str = "" if cell is None else str(cell).strip()
        try:
            if dtype == "Number":
                c_num = float(re.sub(r"[^\d.-]", "", cell_str))
                v_num = float(re.sub(r"[^\d.-]", "", val_str))
                if op == "Less Than" and not (c_num < v_num):
                    return False
                if op == "Greater Than" and not (c_num > v_num):
                    return False
                if op == "Equals" and not (c_num == v_num):
                    return False
                if op == "Contains" and val_str.lower() not in cell_str.lower():
                    return False
            else:
                if op == "Less Than" and not (cell_str.lower() < val_str.lower()):
                    return False
                if op == "Greater Than" and not (cell_str.lower() > val_str.lower()):
                    return False
                if op == "Equals" and not (cell_str.lower() == val_str.lower()):
                    return False
                if op == "Contains" and val_str.lower() not in cell_str.lower():
                    return False
        except Exception:
            return False
    return True


# ---------------------------------------------------------------------------
# Agentic AI Scraper Engine (Gemini 3.8 Flash / Antigravity Agent / Gemini 4 Argon)
# ---------------------------------------------------------------------------

# Fetches run in two steps: discovery calls list many candidate Main-column values at once from pages
# that list them, then parallel fill calls research a few specific candidates each. Parallel calls never
# chase the same entities, and no single response has to carry hundreds of rows.
FETCH_DISCOVERY_SIZE = max(1, int(os.getenv("FETCH_DISCOVERY_SIZE", "30")))
FETCH_FILL_SIZE = max(1, int(os.getenv("FETCH_FILL_SIZE", "4")))
FETCH_CONCURRENCY = max(1, int(os.getenv("FETCH_CONCURRENCY", "6")))
# Fetches run as background jobs split into chunks; each chunk must finish inside the task
# function timeout in main.py (1800s), then saves its state and queues the next chunk.
FETCH_CHUNK_SEC = float(os.getenv("FETCH_CHUNK_SEC", "1500"))
FETCH_JOB_MAX_SEC = float(os.getenv("FETCH_JOB_MAX_HOURS", "6")) * 3600
# A running chunk renews its lease on every save; an expired lease means the chunk died.
FETCH_LEASE_SEC = 180
# Consecutive discoveries yielding under 20% new candidates before the direct query counts as exhausted.
FETCH_MAX_DRY_DISCOVERIES = 3
FETCH_MAX_ERROR_CALLS = 4
FETCH_EXCLUDE_LIST_LIMIT = 400
FETCH_SEEDS_PER_ANGLE = 5
FETCH_JOBS_COLLECTION = "fetch_jobs"

_DIRECT_SEARCH_ANGLES = [
    "",
    "Prioritize lesser-known entries that a quick search would miss.",
    "Prioritize curated lists, directories, and databases over top search results.",
    "Prioritize recent entries (the last 1-2 years).",
    "Prioritize entries from different regions or sub-categories than the obvious ones.",
    "Prioritize official sources (organization websites, institution pages, public registries).",
]

_STOP_REASON_TEXT = {
    "target_reached": "Requested row count reached.",
    "exhausted": "Search exhausted: the direct query and the Main-column expansion searches stopped finding new verifiable records.",
    "time_limit": f"Stopped at the {FETCH_JOB_MAX_SEC / 3600:g}-hour limit for one fetch.",
    "cancelled": "Stopped by request.",
    "error": "Stopped because of an unexpected error.",
    "api_error": "Stopped because the AI provider kept returning errors.",
    "no_api_key": "No Gemini API key configured, so no data could be fetched.",
}


def _main_columns(columns: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return [c for c in columns if c.get("query_type") == "Main" and c.get("name")] or columns[:1]


def _entity_values(rows: List[Dict[str, Any]], col_name: str) -> List[str]:
    """Unique, short, non-empty values of one column in row order."""
    seen: set[str] = set()
    out: List[str] = []
    for r in rows:
        v = str(r.get(col_name, "") or "").strip()
        key = _normalize_entity_str(v)
        if v and len(v) < 60 and key not in seen:
            seen.add(key)
            out.append(v)
    return out


def _build_expansion_angles(
    columns: List[Dict[str, Any]],
    profiles: Dict[str, Dict[str, Any]],
    known_rows: List[Dict[str, Any]],
    used_keys: set[str],
) -> List[str]:
    """
    Search angles used once the direct query runs dry: first new values for each Main column
    seeded from the entities already known, then records matching observed Attribute values.
    """
    main_cols = _main_columns(columns)
    angles: List[str] = []
    for col in main_cols:
        name = col["name"]
        instr = (col.get("special_instructions") or "").strip()
        rules = f" Column rules: {instr}" if instr else ""
        seeds = [s for s in _entity_values(known_rows, name) if f"seed:{name}:{_normalize_entity_str(s)}" not in used_keys]
        for i in range(0, len(seeds), FETCH_SEEDS_PER_ANGLE):
            group = seeds[i:i + FETCH_SEEDS_PER_ANGLE]
            used_keys.update(f"seed:{name}:{_normalize_entity_str(s)}" for s in group)
            angles.append(
                f'Expand on the Main column "{name}": find NEW "{name}" entries closely related to these known ones '
                f"(peers, competitors, same category, or listed alongside them in the same sources): "
                f"{', '.join(group)}.{rules}"
            )
        generic_key = f"generic:{name}"
        if generic_key not in used_keys:
            used_keys.add(generic_key)
            angles.append(
                f'Expand on the Main column "{name}": search specifically for new "{name}" entries using different '
                f"search terms, sources, and listings than a generic query would.{rules}"
            )

    main_names = {c["name"] for c in main_cols}
    for col in columns:
        name = col.get("name", "")
        if not name or name in main_names:
            continue
        for v in (profiles.get(name, {}).get("top_values") or [])[:4]:
            v_str = str(v).strip()
            key = f"attr:{name}:{_normalize_entity_str(v_str)}"
            if v_str and len(v_str) < 60 and key not in used_keys:
                used_keys.add(key)
                angles.append(f'Find new records whose "{name}" is "{v_str}".')
    return angles


def _discovery_schema(main_col_name: str) -> Dict[str, Any]:
    return {
        "type": "object",
        "properties": {
            "candidates": {
                "type": "array",
                "description": f"New, distinct `{main_col_name}` entries found listed on real pages. May be empty.",
                "items": {
                    "type": "object",
                    "properties": {
                        "name": {
                            "type": "string",
                            "description": f"The `{main_col_name}` value exactly as written on the source page.",
                        },
                        "source_url": {
                            "type": "string",
                            "description": "URL of the page where this entry is listed.",
                        },
                    },
                    "required": ["name", "source_url"],
                },
            },
            "agent_summary": {
                "type": "string",
                "description": "Brief summary of where the agent looked.",
            },
            "sources_consulted": {
                "type": "array",
                "items": {"type": "string"},
                "description": "List of web sources, domains, or MCP channels used.",
            },
        },
        "required": ["candidates", "agent_summary", "sources_consulted"],
    }


def _discovery_prompt(size: int, angle: str, exclude: List[str], main_col_name: str, source_names: str) -> str:
    return f"""
Already Known (DO NOT RETURN ANY OF THESE):
{', '.join(exclude) if exclude else 'None'}

Task for this step: discover candidates only.
1. Search the web and prioritized sources ({source_names}) for up to {size} NEW, distinct "{main_col_name}" entries that fit the goal above.
2. Prefer pages that list many entries at once: lists, directories, rosters, rankings, databases.
{f"3. Search focus: {angle}" if angle else "3. Search focus: the user's scraping goal above."}
4. Only return entries that actually appear on a page you found, with that page's URL in `source_url`. Never invent entries.
5. Do not research the other columns yet; return only the entries and where they are listed.
6. Returning fewer than {size} entries, or zero, is correct when no more exist.
"""


def _fill_prompt(candidates: List[Dict[str, str]], main_col_name: str, source_names: str) -> str:
    listing = "\n".join(f"- {c['name']} (listed at {c['source_url']})" for c in candidates)
    return f"""
Task for this step: fill in records for exactly these "{main_col_name}" entries:
{listing}

1. Research each entry, starting with the page it is listed at and then the web and prioritized sources ({source_names}), and fill in every column.
2. Return at most one record per entry, with `{main_col_name}` written exactly as given above. Do not add any other entities.
3. Put the URL of the page that best verifies the record in `_source_url`.
4. If you cannot verify a column's value from a source, return an empty string for it. Never guess, estimate, or invent names, values, or URLs.
5. For `_links`, only include URLs you actually found; use an empty string otherwise. Never construct a URL from a name.
6. Leave an entry out entirely if it turns out not to fit the goal or the column rules.
7. Format every cell to match the Google Sheet conventions above. Return JSON matching the schema.
"""


def _match_candidate(value: Any, by_key: Dict[str, Dict[str, str]]) -> Optional[Dict[str, str]]:
    """The requested candidate a filled row belongs to, tolerating small spelling changes."""
    key = _normalize_entity_str(value)
    if not key:
        return None
    if key in by_key:
        return by_key[key]
    core = _normalize_name_core(value)
    for cand_key, cand in by_key.items():
        cand_core = _normalize_name_core(cand["name"])
        if core and cand_core and len(min(core, cand_core, key=len)) > 3 and (core in cand_core or cand_core in core):
            return cand
    return None


def _run_fetch_batch(
    client: "genai.Client",
    agent_mode: str,
    prompt: str,
    output_json_schema: Dict[str, Any],
    tools_config: List[Dict[str, Any]],
) -> Tuple[Dict[str, Any], str]:
    """One blocking model call for a single small batch. Run via asyncio.to_thread."""
    if agent_mode == "antigravity-agent":
        agent_interaction = client.interactions.create(
            agent="antigravity-preview-09-2026",
            input=prompt + "\nOutput ONLY valid JSON matching the schema.",
            environment="remote",
            tools=[{"type": "google_search"}, {"type": "url_context"}],
        )
        raw_text = agent_interaction.output_text or ""
        struct_interaction = client.interactions.create(
            model="gemini-3.8-flash",
            input=(
                "Convert the following agent research findings into the exact JSON schema. "
                "Do not add any record, value, or URL that is not in the findings:\n\n" + raw_text
            ),
            response_format={
                "type": "text",
                "mime_type": "application/json",
                "schema": output_json_schema,
            },
        )
        return json.loads(struct_interaction.output_text), "antigravity-preview-09-2026 + gemini-3.8-flash"

    target_model = "gemini-4-argon" if agent_mode == "gemini-4-argon" else "gemini-3.8-flash"
    interaction = client.interactions.create(
        model=target_model,
        input=prompt,
        tools=tools_config,
        response_format={
            "type": "text",
            "mime_type": "application/json",
            "schema": output_json_schema,
        },
    )
    return json.loads(interaction.output_text), target_model


def _expand_more_rows(
    row: Dict[str, Any],
    columns: List[Dict[str, Any]],
    profiles: Dict[str, Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """
    Splits bullet-list cells of `retrieve_all_mode == "more_rows"` columns into a vertically merged
    block: row 0 carries every column, continuation rows carry only the next list item.
    """
    more_cols = [
        c["name"] for c in columns
        if c.get("retrieve_all") and c.get("retrieve_all_mode") == "more_rows" and c.get("name")
    ]
    if not more_cols:
        return [row]
    items = {
        n: [ln.strip().lstrip("•-*").strip() for ln in str(row.get(n, "") or "").split("\n") if ln.strip()]
        for n in more_cols
    }
    span = max((len(v) for v in items.values()), default=0)
    if span <= 1:
        return [row]

    group_id = f"grp-{uuid.uuid4().hex[:8]}"
    block: List[Dict[str, Any]] = []
    for i in range(span):
        sub: Dict[str, Any] = {"_group_id": group_id, "_group_span": span, "_group_index": i}
        for c in columns:
            n = c.get("name", "")
            if n in items:
                sub[n] = items[n][i] if i < len(items[n]) else ""
            else:
                sub[n] = row.get(n, "") if i == 0 else ""
        if i == 0:
            for meta_key in ("_links", "_source"):
                if row.get(meta_key):
                    sub[meta_key] = row[meta_key]
        block.append(format_row_for_gsheets(sub, columns, profiles))
    return block


async def _prepare_fetch_context(
    project: Dict[str, Any],
    target_tab: Dict[str, Any],
    query_prompt: str,
    prior_new_rows: List[Dict[str, Any]],
) -> Dict[str, Any]:
    """
    Everything a fetch needs that is derived from the project: aligned columns, learned formats,
    rows to deduplicate against, response schemas, and the shared prompt context.
    """
    all_tabs: List[Dict[str, Any]] = project.get("tabs", [])

    # If the currently active tab is a RULES! tab, use the primary data columns and first data tab's structure
    is_rules_target = target_tab.get("name", "").strip().upper().startswith("RULE")
    effective_tab = target_tab
    if is_rules_target:
        for t in all_tabs:
            if not t.get("name", "").strip().upper().startswith("RULE"):
                effective_tab = t
                break

    columns: List[Dict[str, Any]] = project.get("columns", [])
    tab_headers: List[str] = effective_tab.get("headers", [])

    # Align columns with the Google Sheet tab's exact headers and column order
    if tab_headers and not is_rules_target:
        col_map = {c.get("name", "").strip().upper(): c for c in columns}
        aligned_columns = []
        for idx, h in enumerate(tab_headers):
            existing_c = col_map.get(h.strip().upper())
            if existing_c:
                aligned_columns.append(existing_c)
            else:
                aligned_columns.append({
                    "name": h,
                    "data_type": "String",
                    "query_type": "Main" if idx == 0 else "Attribute",
                    "special_instructions": "",
                    "sort": None,
                    "filter_op": "None",
                    "filter_value": "",
                })
        columns = aligned_columns
    elif not columns:
        columns = [
            {"name": "Name", "data_type": "String", "query_type": "Main", "special_instructions": ""},
            {"name": "Job", "data_type": "String", "query_type": "Attribute", "special_instructions": ""},
            {"name": "Year", "data_type": "Number", "query_type": "Attribute", "special_instructions": ""},
            {"name": "Location", "data_type": "String", "query_type": "Attribute", "special_instructions": ""},
        ]

    # Collect ALL existing rows across all tabs + previously generated rows so we NEVER repeat any existing data
    all_existing_rows: List[Dict[str, Any]] = []
    for t in all_tabs:
        if not t.get("name", "").strip().upper().startswith("RULE"):
            all_existing_rows.extend(t.get("rows", []))
    all_existing_rows.extend(prior_new_rows)

    profiles = learn_sheet_format_profile(effective_tab, all_tabs, columns)

    mcps: List[Dict[str, Any]] = project.get("mcps", [])
    enabled_mcps = [m for m in mcps if m.get("enabled", True)]
    mcp_names = [m.get("name", "") for m in enabled_mcps]

    existing_source = project.get("existing_data_source", "")
    tab_existing_rows = effective_tab.get("rows", [])
    sample_existing_rows = [
        r for r in tab_existing_rows
        if any(str(v).strip() for v in r.values())
    ][:8]

    main_col_name = _main_columns(columns)[0].get("name", "Name") if columns else "Name"

    row_properties: Dict[str, Any] = {}
    required_cols: List[str] = []
    column_descriptions: List[str] = []
    link_properties: Dict[str, Any] = {}

    for col in columns:
        c_name = col.get("name", "Column")
        c_dtype = col.get("data_type", "String")
        c_qtype = col.get("query_type", "Attribute")
        c_instr = col.get("special_instructions", "")
        c_sort = col.get("sort")
        c_filter_op = col.get("filter_op", "None")
        c_filter_val = col.get("filter_value", "")
        prof = profiles.get(c_name, {})

        desc_parts = [f"Query Type: {c_qtype}", f"Data Type: {c_dtype}"]
        if c_instr:
            desc_parts.append(f"Special Instructions: {c_instr}")
        if col.get("put_link"):
            desc_parts.append("Put Link: Enabled (put the real source URL for this value in `_links`)")
            link_properties[c_name] = {
                "type": "string",
                "description": f"Real URL for the `{c_name}` value as found in sources; empty string if none was found.",
            }
        if col.get("retrieve_all"):
            mode_label = (
                "list every item as bullet lines (• Item 1\\n• Item 2); each item becomes its own merged row"
                if col.get("retrieve_all_mode") == "more_rows"
                else "Bullet Points (• Item 1\\n• Item 2)"
            )
            desc_parts.append(f"Retrieve All Sector History: Enabled ({mode_label})")
        if prof.get("top_values"):
            desc_parts.append(f"Observed Sheet Format / Values: {', '.join(str(x) for x in prof['top_values'][:6])}")
        if c_filter_op != "None" and c_filter_val:
            desc_parts.append(f"Must satisfy filter: {c_filter_op} '{c_filter_val}'")
        if c_sort:
            desc_parts.append(f"Sort preference: {c_sort}")
        desc_parts.append("Empty string if not verifiable")

        full_desc = " | ".join(desc_parts)
        # Always a string so an unverifiable Number can be left blank instead of forced to a made-up value;
        # sanitize_cell_for_gsheets converts numeric strings back to numbers.
        row_properties[c_name] = {
            "type": "string",
            "description": full_desc,
        }
        required_cols.append(c_name)
        column_descriptions.append(f"- `{c_name}` ({full_desc})")

    row_properties["_source_url"] = {
        "type": "string",
        "description": "URL of the page where this record was found and verified. Required.",
    }
    required_cols.append("_source_url")
    if link_properties:
        row_properties["_links"] = {"type": "object", "properties": link_properties}

    output_json_schema = {
        "type": "object",
        "properties": {
            "rows": {
                "type": "array",
                "description": "Brand-new, non-duplicate, source-verified records formatted identically to the Google Sheet. May be empty.",
                "items": {
                    "type": "object",
                    "properties": row_properties,
                    "required": required_cols,
                },
            },
            "agent_summary": {
                "type": "string",
                "description": "Brief summary of how the agent searched, verified, and formatted this batch.",
            },
            "sources_consulted": {
                "type": "array",
                "items": {"type": "string"},
                "description": "List of web sources, domains, or MCP channels used.",
            },
        },
        "required": ["rows", "agent_summary", "sources_consulted"],
    }

    combined_text = f"{query_prompt} {existing_source}"
    found_urls = re.findall(r"https?://[^\s,;\")]+", combined_text)
    scraped_web_context = ""
    if found_urls:
        scraped_web_context = await scrape_urls_context(found_urls)

    source_names = ', '.join(mcp_names) if mcp_names else 'Web'
    base_prompt = f"""You are an enterprise Agentic AI Data Mining & Web Scraping specialist. Accuracy matters more than volume: only return real, verifiable data.
Project Name: "{project.get('name', 'Data Mining Project')}"
Target Google Sheet Tab: "{effective_tab.get('name', 'Tab 1')}"
User Scraping Goal / Prompt: "{query_prompt or f'Find real records for {project.get("name", "this project")} ({effective_tab.get("name", "Tab 1")}) matching the exact Google Sheet format.'}"
Active MCP / Source Connectors prioritized: {', '.join(mcp_names) if mcp_names else 'General Web Search'}

Existing Sample Rows Seen in Google Sheet "{effective_tab.get('name', 'Tab 1')}" (MATCH THIS EXACT COLUMN FORMAT, STYLE, AND VALUE CONVENTIONS):
{json.dumps(sample_existing_rows[:6], indent=2)}

Data Template Columns & Learned Google Sheet Formatting Rules:
{chr(10).join(column_descriptions)}

{f"Direct Web Page Context Scraped from Target URLs:{chr(10)}{scraped_web_context}" if scraped_web_context else ""}
"""

    tools_config: List[Dict[str, Any]] = [
        {"type": "google_search"},
        {"type": "url_context"},
    ]
    for m in enabled_mcps:
        srv_url = (m.get("server_url") or "").strip()
        if srv_url.startswith("http"):
            tools_config.append({
                "type": "mcp_server",
                "name": re.sub(r"[^a-zA-Z0-9_-]", "_", m.get("name", "custom_mcp").lower()),
                "url": srv_url,
            })

    return {
        "columns": columns,
        "profiles": profiles,
        "effective_tab": effective_tab,
        "existing_rows": all_existing_rows,
        "main_col_name": main_col_name,
        "row_schema": output_json_schema,
        "discovery_schema": _discovery_schema(main_col_name),
        "base_prompt": base_prompt,
        "source_names": source_names,
        "mcp_names": mcp_names,
        "tools_config": tools_config,
        "headers": [c.get("name", "") for c in columns if c.get("name")],
    }


def _new_fetch_state() -> Dict[str, Any]:
    """Resumable pipeline state, saved on the job document between background chunks."""
    return {
        "phase": "direct",
        "direct_angle_idx": 0,
        "expansion_angles": [],
        "used_expansion_keys": [],
        "candidates": [],
        "seen_candidate_keys": [],
        "dry_discoveries": 0,
        "error_calls": 0,
        "calls": 0,
        "sources_consulted": [],
        "model_used": "",
        "last_error": "",
    }


async def run_fetch_pipeline(
    ctx: Dict[str, Any],
    state: Dict[str, Any],
    rows: List[Dict[str, Any]],
    requested_count: int,
    agent_mode: str,
    api_key: str,
    deadline: float,
    report: Callable[[str], Awaitable[None]],
    should_stop: Callable[[], bool],
) -> str:
    """
    Fills `rows` with real, source-backed records until `requested_count` is reached. Discovery calls
    queue candidate Main-column values; fill calls research FETCH_FILL_SIZE queued candidates each, in
    parallel. When discovery runs dry it expands from the Main columns' known values. `state` and
    `rows` are mutated in place so the caller can save them and resume in a later chunk.

    Returns a stop reason, or "chunk_time" when the loop-clock `deadline` passes first.
    """
    columns = ctx["columns"]
    profiles = ctx["profiles"]
    main_col = ctx["main_col_name"]
    existing_rows = ctx["existing_rows"]

    dedup = DeduplicationEngine(existing_rows, columns)
    for r in rows:
        dedup.register_existing(r)
    seen_keys = set(state["seen_candidate_keys"])
    for v in _entity_values(existing_rows + rows, main_col):
        seen_keys.add(_normalize_entity_str(v))
    used_expansion_keys = set(state["used_expansion_keys"])
    queue: List[Dict[str, str]] = state["candidates"]

    def _accept(raw: Dict[str, Any], fallback_source: str) -> Tuple[List[Dict[str, Any]], str]:
        """Returns the accepted row block, or an empty list and the reason it was skipped."""
        source_url = str(raw.get("_source_url") or "").strip()
        if not source_url.startswith("http"):
            source_url = fallback_source
        formatted_r = format_row_for_gsheets(raw, columns, profiles)
        if not str(formatted_r.get(main_col, "") or "").strip():
            return [], f"no {main_col}"
        raw_links = raw.get("_links") if isinstance(raw.get("_links"), dict) else {}
        links_for_r = {
            c["name"]: str(raw_links.get(c["name"], "")).strip()
            for c in columns
            if c.get("put_link")
            and str(formatted_r.get(c["name"], "")).strip()
            and str(raw_links.get(c["name"], "")).strip().startswith("http")
        }
        if links_for_r:
            formatted_r["_links"] = links_for_r
        formatted_r["_source"] = source_url
        if not _row_matches_column_filters(formatted_r, columns):
            return [], "failed filter"
        if not dedup.accept(formatted_r):
            return [], "duplicate"
        block = _expand_more_rows(formatted_r, columns, profiles)
        for cont_r in block[1:]:
            dedup.register_existing(cont_r)
        return block, ""

    async def _next_discovery_angle() -> Optional[str]:
        if state["phase"] == "direct":
            angle = _DIRECT_SEARCH_ANGLES[state["direct_angle_idx"] % len(_DIRECT_SEARCH_ANGLES)]
            state["direct_angle_idx"] += 1
            return angle
        if not state["expansion_angles"]:
            state["expansion_angles"] = _build_expansion_angles(
                columns, profiles, rows + existing_rows, used_expansion_keys
            )
            state["used_expansion_keys"] = sorted(used_expansion_keys)
            if state["expansion_angles"]:
                await report(
                    f"Expanding the search from Main-column values ({len(state['expansion_angles'])} new search angles)."
                )
        return state["expansion_angles"].pop(0) if state["expansion_angles"] else None

    client = genai.Client(api_key=api_key)
    loop = asyncio.get_running_loop()
    # A dedicated executor so calls still running when we stop never hold up the response
    # (asyncio.run waits on the default executor at shutdown).
    executor = ThreadPoolExecutor(max_workers=FETCH_CONCURRENCY)
    in_flight: Dict["asyncio.Future[Tuple[Dict[str, Any], str]]", Tuple[str, Any]] = {}
    stop_reason = ""

    def _launch(kind: str, payload: Any, task_prompt: str, schema: Dict[str, Any]) -> None:
        fut = loop.run_in_executor(
            executor,
            _run_fetch_batch,
            client,
            agent_mode,
            ctx["base_prompt"] + task_prompt,
            schema,
            ctx["tools_config"],
        )
        in_flight[fut] = (kind, payload)

    try:
        while True:
            remaining = requested_count - len(rows)
            if remaining <= 0:
                stop_reason = "target_reached"
                break
            if should_stop():
                stop_reason = "cancelled"
                break
            if loop.time() >= deadline:
                stop_reason = "chunk_time"
                break

            filling = [c for kind, p in in_flight.values() if kind == "fill" for c in p]
            discovering = sum(1 for kind, _ in in_flight.values() if kind == "discover")

            # Fill queued candidates, assuming roughly 3 in 4 verify and pass the duplicate check.
            while len(in_flight) < FETCH_CONCURRENCY and queue and len(filling) * 0.75 < remaining:
                batch = queue[:FETCH_FILL_SIZE]
                del queue[:FETCH_FILL_SIZE]
                filling.extend(batch)
                _launch("fill", batch, _fill_prompt(batch, main_col, ctx["source_names"]), ctx["row_schema"])

            # Keep a buffer of candidates so fill calls never wait on discovery.
            wants_candidates = len(queue) + len(filling) < remaining * 1.5 + FETCH_FILL_SIZE
            if len(in_flight) < FETCH_CONCURRENCY and wants_candidates and discovering < (1 if queue else 2):
                angle = await _next_discovery_angle()
                if angle is not None:
                    exclude = (
                        [c["name"] for c in queue + filling]
                        + _entity_values(rows[::-1] + existing_rows, main_col)
                    )[:FETCH_EXCLUDE_LIST_LIMIT]
                    _launch(
                        "discover",
                        angle,
                        _discovery_prompt(FETCH_DISCOVERY_SIZE, angle, exclude, main_col, ctx["source_names"]),
                        ctx["discovery_schema"],
                    )

            if not in_flight:
                stop_reason = "exhausted"
                break

            # Wake up at least every few seconds to notice a stop request or the deadline.
            done, _ = await asyncio.wait(
                list(in_flight),
                timeout=max(0.5, min(deadline - loop.time(), 5.0)),
                return_when=asyncio.FIRST_COMPLETED,
            )

            for fut in done:
                kind, payload = in_flight.pop(fut)
                exc = fut.exception()
                if exc is not None:
                    state["error_calls"] += 1
                    state["last_error"] = str(exc)[:120]
                    if kind == "fill":
                        queue[:0] = payload
                    await report(f"A {'discovery' if kind == 'discover' else 'research'} call failed: {state['last_error']}")
                    continue
                state["error_calls"] = 0
                parsed, state["model_used"] = fut.result()
                state["calls"] += 1
                for src in parsed.get("sources_consulted") or []:
                    if src and src not in state["sources_consulted"]:
                        state["sources_consulted"].append(src)
                agent_note = str(parsed.get("agent_summary") or "").strip()

                if kind == "discover":
                    added = 0
                    for cand in parsed.get("candidates") or []:
                        if not isinstance(cand, dict):
                            continue
                        name = str(cand.get("name") or "").strip()
                        source_url = str(cand.get("source_url") or "").strip()
                        key = _normalize_entity_str(name)
                        if not key or not source_url.startswith("http") or key in seen_keys:
                            continue
                        if dedup.is_duplicate({main_col: name}):
                            continue
                        seen_keys.add(key)
                        state["seen_candidate_keys"].append(key)
                        queue.append({"name": name, "source_url": source_url})
                        added += 1
                    msg = "Discovery"
                    if payload:
                        msg += f" ({payload[:90]}{'…' if len(payload) > 90 else ''})"
                    msg += f": {added} new candidates queued, {len(queue)} waiting."
                    if agent_note:
                        msg += f" Agent: {agent_note[:220]}"
                    await report(msg)
                    if state["phase"] == "direct":
                        state["dry_discoveries"] = (
                            state["dry_discoveries"] + 1 if added * 5 < FETCH_DISCOVERY_SIZE else 0
                        )
                        if state["dry_discoveries"] >= FETCH_MAX_DRY_DISCOVERIES:
                            state["phase"] = "expand"
                            await report("Discovery is running dry; switching to Main-column expansion.")
                else:
                    by_key = {_normalize_entity_str(c["name"]): c for c in payload}
                    verified: List[str] = []
                    skipped: Dict[str, int] = {}
                    for raw in parsed.get("rows") or []:
                        if len(rows) >= requested_count:
                            break
                        cand = _match_candidate(raw.get(main_col, ""), by_key) if isinstance(raw, dict) else None
                        if cand is None:
                            skipped["not requested"] = skipped.get("not requested", 0) + 1
                            continue
                        block, reason = _accept(raw, cand["source_url"])
                        if block:
                            rows.extend(block)
                            verified.append(str(block[0].get(main_col, "")))
                        else:
                            skipped[reason] = skipped.get(reason, 0) + 1
                    msg = f"Researched {len(payload)}: {', '.join(verified) if verified else 'none verified'}"
                    if skipped:
                        msg += " · skipped " + ", ".join(f"{n} {r}" for r, n in skipped.items())
                    msg += f". Total {min(len(rows), requested_count)}/{requested_count}."
                    await report(msg)

            if state["error_calls"] >= FETCH_MAX_ERROR_CALLS:
                stop_reason = "api_error"
                break
    finally:
        # Candidates still being researched go back on the queue so a later chunk retries them.
        for fut, (kind, payload) in in_flight.items():
            fut.cancel()
            if kind == "fill":
                queue[:0] = payload
        executor.shutdown(wait=False, cancel_futures=True)
    return stop_reason


# ---------------------------------------------------------------------------
# Background Fetch Jobs (Cloud Tasks chunks + completion email)
# ---------------------------------------------------------------------------

_tasks_app: Optional[firebase_admin.App] = None
_local_job_tasks: set["asyncio.Task[None]"] = set()


def _enqueue_fetch_chunk_sync(job_id: str) -> None:
    global _tasks_app
    # Enqueue with the runtime's own credentials: the deployed function's service account can
    # create Cloud Tasks, while the bundled Admin SDK key may not have that role.
    if _tasks_app is None:
        project_id = os.getenv("GCLOUD_PROJECT") or os.getenv("GOOGLE_CLOUD_PROJECT") or firebase_admin.get_app().project_id
        _tasks_app = firebase_admin.initialize_app(
            credentials.ApplicationDefault(), {"projectId": project_id}, name="tasks"
        )
    queue = admin_functions.task_queue("locations/us-central1/functions/fetchjob", app=_tasks_app)
    queue.enqueue({"job_id": job_id})


async def _enqueue_fetch_chunk(job_id: str) -> None:
    """Runs the next chunk of a fetch job in the background."""
    if os.getenv("K_SERVICE"):
        await asyncio.to_thread(_enqueue_fetch_chunk_sync, job_id)
    else:
        # Local dev (uvicorn): no Cloud Tasks, so run in this process.
        task = asyncio.get_running_loop().create_task(run_fetch_job_chunk(job_id))
        _local_job_tasks.add(task)
        task.add_done_callback(_local_job_tasks.discard)


def _claim_fetch_job(job_ref: Any) -> str:
    """Takes the job's lease. Returns "claimed", "busy" (another chunk holds it), or "finished"."""
    transaction = db.transaction()

    @firestore.transactional
    def _claim(txn: Any) -> str:
        snap = job_ref.get(transaction=txn)
        if not snap.exists:
            return "finished"
        job = snap.to_dict()
        if job.get("status") not in ("queued", "running"):
            return "finished"
        if float(job.get("lease_until") or 0) > time.time():
            return "busy"
        txn.update(job_ref, {
            "status": "running",
            "lease_until": time.time() + FETCH_LEASE_SEC,
            "chunks": int(job.get("chunks") or 0) + 1,
        })
        return "claimed"

    return _claim(transaction)


def _send_fetch_email(to_addr: str, subject: str, body: str) -> str:
    """Sends the completion email from the notification Gmail account. Returns a status note."""
    sender = os.getenv("NOTIFY_EMAIL_FROM", "besabreak@gmail.com").strip()
    password = os.getenv("NOTIFY_EMAIL_APP_PASSWORD", "").replace(" ", "")
    if not to_addr:
        return "not sent: the account has no email address"
    if not password:
        return "not sent: NOTIFY_EMAIL_APP_PASSWORD is not configured"
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f"Data Retrieval System <{sender}>"
    msg["To"] = to_addr
    msg.set_content(body)
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
            smtp.login(sender, password)
            smtp.send_message(msg)
        return f"sent to {to_addr}"
    except Exception as exc:
        return f"not sent: {str(exc)[:150]}"


async def _finalize_fetch_job(
    job_ref: Any,
    job: Dict[str, Any],
    columns: List[Dict[str, Any]],
    rows: List[Dict[str, Any]],
    progress: Dict[str, Any],
    stop_reason: str,
    send_email: bool,
) -> None:
    state = job.get("state") or _new_fetch_state()
    requested = int(job["requested"])
    final_rows = apply_template_transforms(rows, columns)[:requested]
    headers = [c.get("name", "") for c in columns if c.get("name")]

    summary_msg = (
        f"Found {len(final_rows)} of {requested} requested rows, each backed by a source URL "
        f"(0 duplicates against existing sheet rows; {state.get('calls', 0)} AI calls). "
        f"{_STOP_REASON_TEXT.get(stop_reason, '')}"
    )
    if len(final_rows) < requested and stop_reason != "no_api_key":
        summary_msg += " No rows were invented to fill the gap."
    if stop_reason in ("api_error", "error") and state.get("last_error"):
        summary_msg += f" Last error: {state['last_error']}"

    email_status = "not sent: stopped by request"
    if send_email:
        elapsed_min = max(1, round((time.time() - float(job.get("started_epoch") or time.time())) / 60))
        app_url = os.getenv("APP_URL", "").strip()
        body = (
            f"Your data fetch for \"{job.get('project_name', 'your project')}\" has finished.\n\n"
            f"{summary_msg}\n\nTook about {elapsed_min} minute{'s' if elapsed_min != 1 else ''}."
            + (f"\n\nOpen the project: {app_url}" if app_url else "")
        )
        subject = (
            f"Data fetch done: {len(final_rows)}/{requested} rows for {job.get('project_name', 'your project')}"
        )
        email_status = await asyncio.to_thread(_send_fetch_email, job.get("owner_email", ""), subject, body)

    progress["status"] = "done"
    progress["stop_reason"] = stop_reason
    progress["found"] = len(final_rows)
    progress["email_status"] = email_status
    progress["log"] = (
        progress.get("log", [])
        + [{"t": round(time.time() - float(job.get("started_epoch") or time.time()), 1), "msg": summary_msg}]
    )[-40:]
    progress["updated_at"] = _now_iso()

    agent_run_meta = {
        "timestamp": _now_iso(),
        "model_used": state.get("model_used") or job.get("agent_mode", ""),
        "agent_summary": summary_msg,
        "sources_consulted": state.get("sources_consulted") or ["Web Search"],
        "rows_returned": len(final_rows),
        "requested_row_count": requested,
        "deduplicated_against_count": int(job.get("existing_row_count") or 0),
        "target_tab_name": job.get("tab_name", ""),
        "live_ai": len(final_rows) > 0,
        "stop_reason": stop_reason,
        "batches_run": int(state.get("calls") or 0),
        "email_status": email_status,
    }

    proj_ref = db.collection("projects").document(job["project_id"])
    project = (await asyncio.to_thread(proj_ref.get)).to_dict() or {}
    updates: Dict[str, Any] = {
        "new_rows": final_rows,
        "new_headers": headers,
        "last_agent_run": agent_run_meta,
        "fetch_progress": progress,
        "last_edited": _now_iso(),
    }
    tabs = project.get("tabs", [])
    for t in tabs:
        if t.get("id") == job.get("tab_id"):
            if not t.get("headers"):
                t["headers"] = headers
            if job.get("append_rows"):
                t["rows"] = list(t.get("rows", [])) + final_rows
            updates["tabs"] = tabs
            break
    await asyncio.to_thread(proj_ref.update, updates)
    await asyncio.to_thread(job_ref.update, {
        "status": "done",
        "stop_reason": stop_reason,
        "email_status": email_status,
        "finished_at": _now_iso(),
        "lease_until": 0,
        "gemini_api_key": firestore.DELETE_FIELD,
    })


async def run_fetch_job_chunk(job_id: str) -> None:
    """
    Runs one chunk (up to FETCH_CHUNK_SEC) of a background fetch job, saving progress as it goes.
    Queues the next chunk if there is more to do; otherwise finalizes the job and sends the email.
    Raises RuntimeError when another chunk holds the lease, so Cloud Tasks retries later.
    """
    load_dotenv(BASE_DIR / ".env", override=True)
    load_dotenv(BASE_DIR / ".env.example", override=False)

    job_ref = db.collection(FETCH_JOBS_COLLECTION).document(job_id)
    claim = await asyncio.to_thread(_claim_fetch_job, job_ref)
    if claim == "finished":
        return
    if claim == "busy":
        raise RuntimeError(f"Fetch job {job_id} is already running in another chunk")

    job = (await asyncio.to_thread(job_ref.get)).to_dict()
    proj_ref = db.collection("projects").document(job["project_id"])
    project = (await asyncio.to_thread(proj_ref.get)).to_dict() or {}
    progress: Dict[str, Any] = dict(project.get("fetch_progress") or {})
    if progress.get("run_id") != job_id:
        # A newer fetch or a clear replaced this one.
        await asyncio.to_thread(job_ref.update, {
            "status": "superseded", "lease_until": 0, "gemini_api_key": firestore.DELETE_FIELD,
        })
        return

    started_epoch = float(job["started_epoch"])
    requested = int(job["requested"])
    rows: List[Dict[str, Any]] = list(project.get("new_rows") or [])
    state: Dict[str, Any] = {**_new_fetch_state(), **(job.get("state") or {})}
    job["state"] = state
    columns: List[Dict[str, Any]] = project.get("columns", [])
    cancel_requested = bool(job.get("cancel_requested"))
    dirty = asyncio.Event()

    async def _report(message: str) -> None:
        progress["found"] = min(len(rows), requested)
        progress["log"] = (
            progress.get("log", []) + [{"t": round(time.time() - started_epoch, 1), "msg": message}]
        )[-40:]
        dirty.set()

    async def _flush() -> None:
        nonlocal cancel_requested
        dirty.clear()
        progress["updated_at"] = _now_iso()
        await asyncio.to_thread(
            proj_ref.update, {"fetch_progress": copy.deepcopy(progress), "new_rows": list(rows[:requested])}
        )
        await asyncio.to_thread(
            job_ref.update, {"state": copy.deepcopy(state), "lease_until": time.time() + FETCH_LEASE_SEC}
        )
        snap = (await asyncio.to_thread(job_ref.get)).to_dict() or {}
        cancel_requested = cancel_requested or bool(snap.get("cancel_requested"))

    async def _flusher() -> None:
        while True:
            try:
                await asyncio.wait_for(dirty.wait(), timeout=60)
            except asyncio.TimeoutError:
                pass  # Still renew the lease and check for a stop request during long calls.
            await _flush()
            await asyncio.sleep(1.5)

    stop_reason = "error"
    flusher = asyncio.create_task(_flusher())
    try:
        api_key = (job.get("gemini_api_key") or os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or "").strip()
        tabs = project.get("tabs", [])
        target_tab = next(
            (t for t in tabs if t.get("id") == job.get("tab_id")),
            tabs[0] if tabs else {"id": job.get("tab_id"), "name": "Tab 1", "rows": []},
        )
        ctx = await _prepare_fetch_context(project, target_tab, job.get("query_prompt", ""), job.get("prior_new_rows") or [])
        columns = ctx["columns"]
        if int(job.get("chunks") or 0) <= 1:
            await _report(
                f"Searching for {requested} rows with {job.get('agent_mode')}: discovery finds up to "
                f"{FETCH_DISCOVERY_SIZE} candidates per call, then up to {FETCH_CONCURRENCY} calls research "
                f"{FETCH_FILL_SIZE} candidates each."
            )
        else:
            await _report(f"Resuming in background run #{job.get('chunks')}.")

        if not api_key:
            stop_reason = "no_api_key"
        else:
            remaining_job_sec = started_epoch + FETCH_JOB_MAX_SEC - time.time()
            loop = asyncio.get_running_loop()
            stop_reason = await run_fetch_pipeline(
                ctx=ctx,
                state=state,
                rows=rows,
                requested_count=requested,
                agent_mode=job.get("agent_mode", "gemini-3.8-flash"),
                api_key=api_key,
                deadline=loop.time() + max(0.0, min(FETCH_CHUNK_SEC, remaining_job_sec)),
                report=_report,
                should_stop=lambda: cancel_requested,
            )
    except Exception as exc:
        state["last_error"] = str(exc)[:200]
        await _report(f"Fetch failed: {state['last_error']}")
    finally:
        flusher.cancel()

    if stop_reason == "chunk_time" and time.time() < started_epoch + FETCH_JOB_MAX_SEC - 5:
        await _report("Continuing in a new background run.")
        await _flush()
        await asyncio.to_thread(job_ref.update, {"lease_until": 0})
        try:
            await _enqueue_fetch_chunk(job_id)
            return
        except Exception as exc:
            state["last_error"] = f"Could not queue the next background run: {str(exc)[:150]}"
            stop_reason = "error"
    if stop_reason == "chunk_time":
        stop_reason = "time_limit"

    await _flush()
    await _finalize_fetch_job(
        job_ref, job, columns, rows, progress, stop_reason, send_email=stop_reason != "cancelled",
    )


# ---------------------------------------------------------------------------
# API Routes
# ---------------------------------------------------------------------------

@app.get("/api/health")
async def health_check():
    load_dotenv(BASE_DIR / ".env", override=True)
    load_dotenv(BASE_DIR / ".env.example", override=False)
    has_gemini_key = bool((os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or "").strip())
    return {
        "status": "ok",
        "firebase_project": "dr-system-c7488",
        "service_account_email": "firebase-adminsdk-fbsvc@dr-system-c7488.iam.gserviceaccount.com",
        "gemini_api_key_configured": has_gemini_key,
    }


@app.get("/api/projects")
async def list_projects(user: Dict[str, Any] = Depends(get_current_user)):
    uid = user["uid"]
    email = user.get("email", "")

    projects_ref = db.collection("projects")
    query_snap = projects_ref.where("owner_uid", "==", uid).stream()

    projects: List[Dict[str, Any]] = []
    for doc in query_snap:
        data = doc.to_dict()
        data["id"] = doc.id
        projects.append(data)

    if not projects:
        starters = _build_starter_projects(uid, email)
        for starter in starters:
            _, doc_ref = projects_ref.add(starter)
            starter_copy = dict(starter)
            starter_copy["id"] = doc_ref.id
            projects.append(starter_copy)

    projects.sort(key=lambda p: p.get("last_edited", ""), reverse=True)
    return {"projects": projects}


@app.post("/api/projects")
async def create_project(
    payload: ProjectCreateRequest,
    user: Dict[str, Any] = Depends(get_current_user),
):
    now = _now_iso()
    default_headers = ["Name", "Job", "Year", "Location"]
    new_project = {
        "name": payload.name.strip() or "Untitled-project",
        "owner_uid": user["uid"],
        "owner_email": user.get("email", ""),
        "created_at": now,
        "last_edited": now,
        "existing_data_source": "",
        "existing_data_type": "none",
        "active_tab_id": "tab-1",
        "columns": [
            {
                "id": "col-name",
                "name": "Name",
                "special_instructions": "",
                "sort": "Ascending",
                "query_type": "Main",
                "data_type": "String",
                "filter_op": "Less Than",
                "filter_value": "",
            },
            {
                "id": "col-job",
                "name": "Job",
                "special_instructions": "",
                "sort": None,
                "query_type": "Attribute",
                "data_type": "String",
                "filter_op": "None",
                "filter_value": "",
            },
            {
                "id": "col-year",
                "name": "Year",
                "special_instructions": "",
                "sort": None,
                "query_type": "Attribute",
                "data_type": "Number",
                "filter_op": "None",
                "filter_value": "",
            },
            {
                "id": "col-location",
                "name": "Location",
                "special_instructions": "",
                "sort": None,
                "query_type": "Attribute",
                "data_type": "String",
                "filter_op": "None",
                "filter_value": "",
            },
        ],
        "tabs": [
            {"id": "tab-1", "name": "Tab 1", "headers": default_headers, "rows": []},
            {"id": "tab-2", "name": "Tab 2", "headers": default_headers, "rows": []},
        ],
        "mcps": _default_mcps(),
        "last_agent_run": None,
    }
    _, doc_ref = db.collection("projects").add(new_project)
    new_project["id"] = doc_ref.id
    return new_project


@app.get("/api/projects/{project_id}")
async def get_project(
    project_id: str,
    user: Dict[str, Any] = Depends(get_current_user),
):
    doc_ref = db.collection("projects").document(project_id)
    doc = doc_ref.get()
    if not doc.exists:
        raise HTTPException(status_code=404, detail="Project not found")
    data = doc.to_dict()
    if data.get("owner_uid") != user["uid"]:
        raise HTTPException(status_code=403, detail="Not authorized to access this project")
    data["id"] = doc.id
    return data


@app.put("/api/projects/{project_id}")
async def update_project(
    project_id: str,
    payload: ProjectUpdateRequest,
    user: Dict[str, Any] = Depends(get_current_user),
):
    doc_ref = db.collection("projects").document(project_id)
    doc = doc_ref.get()
    if not doc.exists:
        raise HTTPException(status_code=404, detail="Project not found")
    existing = doc.to_dict()
    if existing.get("owner_uid") != user["uid"]:
        raise HTTPException(status_code=403, detail="Not authorized to modify this project")

    updates: Dict[str, Any] = {"last_edited": _now_iso()}
    if payload.name is not None:
        updates["name"] = payload.name.strip() or existing.get("name", "Untitled")
    if payload.existing_data_source is not None:
        updates["existing_data_source"] = payload.existing_data_source
    if payload.existing_data_type is not None:
        updates["existing_data_type"] = payload.existing_data_type
    if payload.active_tab_id is not None:
        updates["active_tab_id"] = payload.active_tab_id
    if payload.tabs is not None:
        updates["tabs"] = [t.model_dump() for t in payload.tabs]
    if payload.columns is not None:
        updates["columns"] = [c.model_dump() for c in payload.columns]
    if payload.mcps is not None:
        updates["mcps"] = [m.model_dump() for m in payload.mcps]
    if payload.new_rows is not None:
        updates["new_rows"] = payload.new_rows

    # If sync_to_gsheet is requested and this project is backed by a Google Sheet, push active tab
    gsheet_sync_status = None
    eff_type = updates.get("existing_data_type", existing.get("existing_data_type"))
    eff_source = updates.get("existing_data_source", existing.get("existing_data_source", ""))
    if payload.sync_to_gsheet and eff_type == "gsheets" and extract_spreadsheet_id(eff_source):
        eff_tabs = updates.get("tabs", existing.get("tabs", []))
        eff_active_id = updates.get("active_tab_id", existing.get("active_tab_id"))
        eff_cols = updates.get("columns", existing.get("columns", []))
        active_tab = next((t for t in eff_tabs if t.get("id") == eff_active_id), None)
        if active_tab:
            gsheet_sync_status = push_tab_to_google_sheet(eff_source, active_tab, eff_cols)
            updates["tabs"] = eff_tabs

    doc_ref.update(updates)
    updated_doc = doc_ref.get().to_dict()
    updated_doc["id"] = project_id
    if gsheet_sync_status:
        updated_doc["gsheet_sync"] = gsheet_sync_status
    return updated_doc


@app.post("/api/projects/{project_id}/cell")
async def update_sheet_cell(
    project_id: str,
    payload: CellUpdateRequest,
    user: Dict[str, Any] = Depends(get_current_user),
):
    """
    Fast endpoint for updating a single cell in a tab and syncing that exact cell to Google Sheets.
    """
    doc_ref = db.collection("projects").document(project_id)
    doc = doc_ref.get()
    if not doc.exists:
        raise HTTPException(status_code=404, detail="Project not found")
    project = doc.to_dict()
    if project.get("owner_uid") != user["uid"]:
        raise HTTPException(status_code=403, detail="Not authorized")

    tabs = project.get("tabs", [])
    target_tab = next((t for t in tabs if t.get("id") == payload.tab_id), None)
    if not target_tab:
        raise HTTPException(status_code=404, detail="Tab not found")

    rows = target_tab.get("rows", [])
    if 0 <= payload.row_index < len(rows):
        rows[payload.row_index][payload.column_name] = payload.value

    gsheet_sync = None
    if (
        payload.sync_to_gsheet
        and project.get("existing_data_type") == "gsheets"
        and extract_spreadsheet_id(project.get("existing_data_source", ""))
    ):
        gsheet_sync = update_single_cell_in_google_sheet(
            project["existing_data_source"],
            target_tab,
            payload.row_index,
            payload.column_name,
            payload.value,
            project.get("columns", []),
        )

    doc_ref.update({"tabs": tabs, "last_edited": _now_iso()})
    return {"ok": True, "gsheet_sync": gsheet_sync}


@app.post("/api/projects/{project_id}/add-row")
async def add_row_to_tab(
    project_id: str,
    payload: AddRowRequest,
    user: Dict[str, Any] = Depends(get_current_user),
):
    doc_ref = db.collection("projects").document(project_id)
    doc = doc_ref.get()
    if not doc.exists:
        raise HTTPException(status_code=404, detail="Project not found")
    project = doc.to_dict()
    if project.get("owner_uid") != user["uid"]:
        raise HTTPException(status_code=403, detail="Not authorized")

    tabs = project.get("tabs", [])
    target_tab = next((t for t in tabs if t.get("id") == payload.tab_id), None)
    if not target_tab:
        raise HTTPException(status_code=404, detail="Tab not found")

    headers = target_tab.get("headers") or [c.get("name", "") for c in project.get("columns", [])]
    new_row = payload.row or {h: "" for h in headers}
    target_tab.setdefault("rows", []).append(new_row)

    gsheet_sync = None
    if (
        payload.sync_to_gsheet
        and project.get("existing_data_type") == "gsheets"
        and extract_spreadsheet_id(project.get("existing_data_source", ""))
    ):
        gsheet_sync = push_tab_to_google_sheet(
            project["existing_data_source"],
            target_tab,
            project.get("columns", []),
            append_only_rows=[new_row],
        )

    doc_ref.update({"tabs": tabs, "last_edited": _now_iso()})
    updated = doc_ref.get().to_dict()
    updated["id"] = project_id
    updated["gsheet_sync"] = gsheet_sync
    return updated


@app.delete("/api/projects/{project_id}")
async def delete_project(
    project_id: str,
    user: Dict[str, Any] = Depends(get_current_user),
):
    doc_ref = db.collection("projects").document(project_id)
    doc = doc_ref.get()
    if not doc.exists:
        raise HTTPException(status_code=404, detail="Project not found")
    existing = doc.to_dict()
    if existing.get("owner_uid") != user["uid"]:
        raise HTTPException(status_code=403, detail="Not authorized to delete this project")
    doc_ref.delete()
    return {"deleted": True, "id": project_id}


@app.post("/api/projects/{project_id}/upload-csv")
async def upload_csv_to_project(
    project_id: str,
    file: UploadFile = File(...),
    tab_id: Optional[str] = Form(default=None),
    sync_columns: bool = Form(default=True),
    user: Dict[str, Any] = Depends(get_current_user),
):
    doc_ref = db.collection("projects").document(project_id)
    doc = doc_ref.get()
    if not doc.exists:
        raise HTTPException(status_code=404, detail="Project not found")
    project = doc.to_dict()
    if project.get("owner_uid") != user["uid"]:
        raise HTTPException(status_code=403, detail="Not authorized")

    raw_bytes = await file.read()
    text = raw_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    headers = [h.strip() for h in (reader.fieldnames or []) if h and h.strip()]
    parsed_rows: List[Dict[str, Any]] = []
    for row in reader:
        clean_row = {k.strip(): (v.strip() if isinstance(v, str) else v) for k, v in row.items() if k and k.strip()}
        if any(clean_row.values()):
            parsed_rows.append(clean_row)

    target_tab_id = tab_id or project.get("active_tab_id", "tab-1")
    tabs = project.get("tabs", [])
    tab_found = False
    for t in tabs:
        if t.get("id") == target_tab_id:
            t["headers"] = headers
            t["rows"] = parsed_rows
            tab_found = True
            break
    if not tab_found:
        tabs.append({"id": target_tab_id, "name": file.filename or "Tab 1", "headers": headers, "rows": parsed_rows})

    columns = project.get("columns", [])
    if sync_columns and headers:
        existing_col_map = {c.get("name", "").lower(): c for c in columns}
        new_columns = []
        for idx, h in enumerate(headers):
            if h.lower() in existing_col_map:
                new_columns.append(existing_col_map[h.lower()])
            else:
                new_columns.append({
                    "id": f"col-{uuid.uuid4().hex[:8]}",
                    "name": h,
                    "special_instructions": "",
                    "sort": "Ascending" if idx == 0 else None,
                    "query_type": "Main" if idx == 0 else "Attribute",
                    "data_type": "String",
                    "filter_op": "None",
                    "filter_value": "",
                })
        columns = new_columns

    updates = {
        "existing_data_source": file.filename or "uploaded.csv",
        "existing_data_type": "csv",
        "tabs": tabs,
        "columns": columns,
        "last_edited": _now_iso(),
    }
    doc_ref.update(updates)
    updated = doc_ref.get().to_dict()
    updated["id"] = project_id
    return updated


@app.post("/api/projects/{project_id}/import-gsheet")
async def import_google_sheet(
    project_id: str,
    payload: GSheetImportRequest,
    user: Dict[str, Any] = Depends(get_current_user),
):
    """
    Imports ALL tabs, rows, columns, and RULES! instructions from a Google Sheet URL.
    """
    doc_ref = db.collection("projects").document(project_id)
    doc = doc_ref.get()
    if not doc.exists:
        raise HTTPException(status_code=404, detail="Project not found")
    project = doc.to_dict()
    if project.get("owner_uid") != user["uid"]:
        raise HTTPException(status_code=403, detail="Not authorized")

    sheet_url = payload.url.strip()
    pulled = await pull_all_tabs_from_google_sheet(sheet_url, project.get("columns", []))

    updates = {
        "existing_data_source": sheet_url,
        "existing_data_type": "gsheets",
        "tabs": pulled["tabs"],
        "columns": pulled["columns"] if payload.sync_columns else project.get("columns", []),
        "active_tab_id": pulled["active_tab_id"],
        "last_edited": _now_iso(),
    }
    doc_ref.update(updates)
    updated = doc_ref.get().to_dict()
    updated["id"] = project_id
    updated["gsheet_sync"] = {
        "synced": True,
        "mode": "pull_all_tabs",
        "spreadsheet_title": pulled["spreadsheet_title"],
        "tab_count": len(pulled["tabs"]),
        "can_write": pulled["can_write"],
    }
    return updated


@app.post("/api/projects/{project_id}/sync-gsheet")
async def sync_google_sheet(
    project_id: str,
    payload: GSheetSyncRequest,
    user: Dict[str, Any] = Depends(get_current_user),
):
    """
    Explicit Two-Way Sync endpoint:
      - direction="pull": Re-fetches all tabs, headers, rows, and rules from the linked Google Sheet.
      - direction="push": Writes the active tab (or all tabs) back to the live Google Sheet.
    """
    doc_ref = db.collection("projects").document(project_id)
    doc = doc_ref.get()
    if not doc.exists:
        raise HTTPException(status_code=404, detail="Project not found")
    project = doc.to_dict()
    if project.get("owner_uid") != user["uid"]:
        raise HTTPException(status_code=403, detail="Not authorized")

    sheet_url = project.get("existing_data_source", "").strip()
    if not extract_spreadsheet_id(sheet_url):
        raise HTTPException(status_code=400, detail="No Google Sheets URL is linked to this project.")

    if payload.direction == "pull":
        pulled = await pull_all_tabs_from_google_sheet(sheet_url, project.get("columns", []))
        current_active = payload.tab_id or project.get("active_tab_id")
        valid_ids = {t["id"] for t in pulled["tabs"]}
        next_active = current_active if current_active in valid_ids else pulled["active_tab_id"]

        updates = {
            "existing_data_type": "gsheets",
            "tabs": pulled["tabs"],
            "columns": pulled["columns"],
            "active_tab_id": next_active,
            "last_edited": _now_iso(),
        }
        doc_ref.update(updates)
        updated = doc_ref.get().to_dict()
        updated["id"] = project_id
        updated["gsheet_sync"] = {
            "synced": True,
            "mode": "pull",
            "spreadsheet_title": pulled["spreadsheet_title"],
            "tab_count": len(pulled["tabs"]),
        }
        return updated

    # direction == "push"
    tabs = project.get("tabs", [])
    columns = project.get("columns", [])
    target_tab_id = payload.tab_id or project.get("active_tab_id")
    results = []
    for t in tabs:
        if payload.tab_id and t.get("id") != target_tab_id:
            continue
        res = push_tab_to_google_sheet(sheet_url, t, columns)
        results.append(res)

    doc_ref.update({"tabs": tabs, "last_edited": _now_iso()})
    updated = doc_ref.get().to_dict()
    updated["id"] = project_id
    updated["gsheet_sync"] = {
        "synced": all(r.get("synced") for r in results) if results else False,
        "mode": "push",
        "details": results,
    }
    return updated


@app.post("/api/projects/{project_id}/fetch-data")
async def fetch_data_with_agent(
    project_id: str,
    payload: FetchDataRequest,
    user: Dict[str, Any] = Depends(get_current_user),
):
    doc_ref = db.collection("projects").document(project_id)
    doc = doc_ref.get()
    if not doc.exists:
        raise HTTPException(status_code=404, detail="Project not found")
    project = doc.to_dict()
    if project.get("owner_uid") != user["uid"]:
        raise HTTPException(status_code=403, detail="Not authorized")

    # If the user pasted a Google Sheets URL in Existing Data without clicking Google Sheets first, pull it automatically
    source_str = (project.get("existing_data_source") or "").strip()
    if "docs.google.com/spreadsheets" in source_str and extract_spreadsheet_id(source_str):
        has_gs_tabs = any(str(t.get("id", "")).startswith("tab-gs-") for t in project.get("tabs", []))
        if not has_gs_tabs or project.get("existing_data_type") != "gsheets":
            try:
                pulled = await pull_all_tabs_from_google_sheet(source_str, project.get("columns", []))
                project["tabs"] = pulled["tabs"]
                project["columns"] = pulled["columns"]
                project["active_tab_id"] = pulled["active_tab_id"]
                project["existing_data_type"] = "gsheets"
            except Exception:
                pass

    target_tab_id = payload.tab_id or project.get("active_tab_id", "tab-1")
    tabs = project.get("tabs", [])
    target_tab = next(
        (t for t in tabs if t.get("id") == target_tab_id),
        tabs[0] if tabs else {"id": target_tab_id, "name": "Tab 1", "rows": []},
    )

    if (project.get("fetch_progress") or {}).get("status") == "running":
        raise HTTPException(status_code=409, detail="A fetch is already running for this project. Stop it first.")

    requested_lines = max(1, min(int(payload.row_count), 1000))
    job_id = payload.run_id or uuid.uuid4().hex[:12]
    started_epoch = time.time()
    prior_new_rows = project.get("new_rows") or []
    existing_row_count = len(prior_new_rows) + sum(
        len(t.get("rows", [])) for t in tabs if not t.get("name", "").strip().upper().startswith("RULE")
    )

    # The fetch runs as a background job; the client polls /fetch-progress and the owner is emailed
    # when it finishes. A new request replaces the previous new data right away.
    job: Dict[str, Any] = {
        "project_id": project_id,
        "project_name": project.get("name", ""),
        "owner_uid": user["uid"],
        "owner_email": user.get("email", ""),
        "tab_id": target_tab.get("id"),
        "tab_name": target_tab.get("name", "Tab 1"),
        "query_prompt": payload.query_prompt or "",
        "requested": requested_lines,
        "agent_mode": payload.agent_mode,
        "append_rows": payload.append_rows,
        # Previous results stay excluded so a new fetch finds different records.
        "prior_new_rows": prior_new_rows,
        "existing_row_count": existing_row_count,
        "state": _new_fetch_state(),
        "status": "queued",
        "lease_until": 0,
        "chunks": 0,
        "cancel_requested": False,
        "started_epoch": started_epoch,
        "created_at": _now_iso(),
    }
    if payload.gemini_api_key:
        # Needed by the background run; deleted from the job when it finishes.
        job["gemini_api_key"] = payload.gemini_api_key
    progress: Dict[str, Any] = {
        "run_id": job_id,
        "status": "running",
        "requested": requested_lines,
        "found": 0,
        "log": [{"t": 0, "msg": "Queued to run in the background."}],
        "notify_email": user.get("email", ""),
        "started_at": _now_iso(),
        "updated_at": _now_iso(),
    }

    job_ref = db.collection(FETCH_JOBS_COLLECTION).document(job_id)
    await asyncio.to_thread(job_ref.set, job)
    await asyncio.to_thread(doc_ref.update, {
        "tabs": tabs,
        "columns": project.get("columns", []),
        "existing_data_type": project.get("existing_data_type", "none"),
        "active_tab_id": target_tab.get("id"),
        "new_rows": [],
        "last_agent_run": None,
        "fetch_progress": progress,
        "last_edited": _now_iso(),
    })
    try:
        await _enqueue_fetch_chunk(job_id)
    except Exception as exc:
        detail = f"Could not start the background fetch: {str(exc)[:200]}"
        progress["status"] = "error"
        progress["log"].append({"t": 0, "msg": detail})
        await asyncio.to_thread(job_ref.update, {"status": "error", "gemini_api_key": firestore.DELETE_FIELD})
        await asyncio.to_thread(doc_ref.update, {"fetch_progress": progress})
        raise HTTPException(status_code=500, detail=detail)

    updated = doc_ref.get().to_dict()
    updated["id"] = project_id
    return updated


@app.post("/api/projects/{project_id}/fetch-cancel")
async def cancel_fetch(
    project_id: str,
    user: Dict[str, Any] = Depends(get_current_user),
):
    """Asks the running fetch to stop; it keeps the rows found so far."""
    doc_ref = db.collection("projects").document(project_id)
    doc = doc_ref.get()
    if not doc.exists:
        raise HTTPException(status_code=404, detail="Project not found")
    project = doc.to_dict()
    if project.get("owner_uid") != user["uid"]:
        raise HTTPException(status_code=403, detail="Not authorized")

    progress = dict(project.get("fetch_progress") or {})
    if progress.get("status") == "running" and progress.get("run_id"):
        job_ref = db.collection(FETCH_JOBS_COLLECTION).document(progress["run_id"])
        job_snap = job_ref.get()
        if job_snap.exists:
            await asyncio.to_thread(job_ref.update, {"cancel_requested": True})
            job = job_snap.to_dict()
            if float(job.get("lease_until") or 0) < time.time():
                # No chunk is running to notice the request (queued or stalled), so stop it here.
                await _finalize_fetch_job(
                    job_ref, job, project.get("columns", []), project.get("new_rows") or [],
                    progress, "cancelled", send_email=False,
                )
        else:
            progress["status"] = "done"
            progress["stop_reason"] = "cancelled"
            await asyncio.to_thread(doc_ref.update, {"fetch_progress": progress})

    updated = doc_ref.get().to_dict()
    updated["id"] = project_id
    return updated


@app.get("/api/projects/{project_id}/fetch-progress")
async def get_fetch_progress(
    project_id: str,
    since: int = 0,
    user: Dict[str, Any] = Depends(get_current_user),
):
    """Live progress of the current fetch plus the new rows found after index `since`."""
    doc = db.collection("projects").document(project_id).get()
    if not doc.exists:
        raise HTTPException(status_code=404, detail="Project not found")
    project = doc.to_dict()
    if project.get("owner_uid") != user["uid"]:
        raise HTTPException(status_code=403, detail="Not authorized")
    rows = project.get("new_rows") or []
    return {
        "progress": project.get("fetch_progress"),
        "total_rows": len(rows),
        "rows": rows[max(0, since):],
    }


@app.delete("/api/projects/{project_id}/new-data")
async def clear_new_data(
    project_id: str,
    user: Dict[str, Any] = Depends(get_current_user),
):
    doc_ref = db.collection("projects").document(project_id)
    doc = doc_ref.get()
    if not doc.exists:
        raise HTTPException(status_code=404, detail="Project not found")
    if doc.to_dict().get("owner_uid") != user["uid"]:
        raise HTTPException(status_code=403, detail="Not authorized")
    if (doc.to_dict().get("fetch_progress") or {}).get("status") == "running":
        raise HTTPException(status_code=409, detail="A fetch is still running. Stop it before clearing.")
    doc_ref.update({
        "new_rows": [],
        "new_headers": [],
        "last_agent_run": None,
        "fetch_progress": None,
        "last_edited": _now_iso(),
    })
    updated = doc_ref.get().to_dict()
    updated["id"] = project_id
    return updated


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
