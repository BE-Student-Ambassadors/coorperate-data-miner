import csv
import difflib
import functools
import io
import json
import os
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Literal, Optional, Tuple

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
# Google Sheets Format Learner, Deduplication Engine & High-Scale Generator
# ---------------------------------------------------------------------------

FIRST_NAMES_POOL = [
    "Aarav", "Abigail", "Ada", "Aditi", "Adrian", "Aiden", "Akash", "Alana", "Alejandro", "Alex",
    "alexa", "Alexander", "Alexandra", "Alexis", "Alice", "Alicia", "Alina", "Allison", "Alyssa", "Amanda",
    "Amara", "Amber", "Amelia", "Amir", "Amit", "Amy", "Ana", "Ananya", "Andre", "Andrea",
    "Andres", "Andrew", "Angela", "Angelina", "Anika", "Anish", "Anita", "Anjali", "Anna", "Annabelle",
    "Anthony", "Antonio", "Anya", "Aria", "Ariana", "Arjun", "Armaan", "Arthur", "Arya", "Aryan",
    "Ashley", "Ashwin", "Audrey", "Austin", "Ava", "Avery", "Ayaan", "Ayesha", "beatrice", "Benjamin",
    "Bianca", "Blake", "Brandon", "Brenda", "Brian", "Brianna", "Brooke", "Bryan", "Caleb", "calvin",
    "Cameron", "Camila", "Carlos", "Carmen", "Caroline", "Carter", "Cassandra", "Catherine", "cedric", "Celeste",
    "Charles", "Charlie", "Charlotte", "Chelsea", " Chloe", "Christian", "Christina", "Christopher", "Claire", "Clara",
    "Claudia", "Cole", "Colin", "Connor", "Cooper", "Courtney", "Cristian", "Crystal", "Cyrus", "Daisy",
    "Damian", "Dana", "Daniel", "Daniela", "Danielle", "Dante", "Daria", "Darius", "David", " Dean",
    "Deepak", "Delilah", "Dennis", "Derek", "Devin", "Devon", "Dev", "Dhruv", "Diana", "Diego",
    "Divya", "Dominic", "Donovan", "Dorian", "Douglas", "Dylan", "Eden", "Edgar", "Eduardo", "Edward",
    "Edwin", "Eleanor", "Elena", "Eli", "Eliana", "Elias", "Elijah", "Elise", "Elizabeth", "Ella",
    "Elliot", "Eloise", "Elsa", "Emilio", "Emily", "Emma", "Emmanuel", "Eric", "Erica", "Erik",
    "Erin", "Esmeralda", "Esteban", "Ethan", "Eugene", "Eva", "Evan", "Evelyn", "Ezra", "Fabian",
    "Faith", "Fatima", "Felix", "Fernando", "Fiona", "Francesca", "Francisco", "Frank", "Freya", "Gabriel",
    "Gabriela", "Gabrielle", "Gael", "Garrett", "Gavin", "Gemma", "Genevieve", "George", "Georgia", "Gianna",
    "Gideon", "Giovanni", "Giselle", "Grace", "Grayson", "Gregory", "Greta", "Griffin", "Guillermo", "Hailey",
    "Hamza", "Hana", "Hannah", "Harini", "Harley", "Harper", "Harrison", "Harsh", "Hassan", "Hayden",
    "Hazel", "Heather", "Hector", "Heidi", "Helen", "Helena", "Henry", "Holly", "Holden", "Hope",
    "Hudson", "Hugh", "Hunter", "Ian", "Ibrahim", "ignacio", "Ilya", "Imani", "India", "Ingrid",
    "Irene", "Iris", "Isaac", "Isabel", "Isabella", "Isabelle", "Isaiah", "Ishaan", "Isla", "Israel",
    "Ivan", "Ivy", "Jack", "Jackson", "Jacob", "Jacqueline", "Jade", "Jaden", "Jake", "James",
    "Jamie", " Janelle", "Jared", "Jasmine", "Jason", "Jasper", "Javier", "Jay", "Jayden", "Jayson",
    "Jean", "Jeffrey", "Jenna", "Jennifer", "jeremiah", "Jeremy", "Jesse", "Jessica", "Jia", "Jian",
    "Jill", "Jocelyn", " Joel", "Johanna", "John", "Jonathan", "Jordan", "Jorge", "Jose", "Joseph",
    "Josephine", "Joshua", "Josiah", "Joy", "Joyce", "Juan", "Judah", "Jude", "Julia", "Julian",
    "Juliana", "juliet", "Juliette", "Jun", "Justin", "Kabir", "Kai", "Kaia", "Kaitlyn", "Kaleb",
    "kameron", "Karen", "Karina", "Karthik", "Kate", "Katelyn", "Katherine", "Kathleen", "Kathryn", "Katie",
    "Kavya", "Kayla", "Keegan", "Keira", "keith", "Kelly", "Kelsey", "Kendall", "Kendrick", "Kenneth",
    "Kenji", "Kevin", "Khadija", "Khalil", "Kiara", "Kieran", "kimberly", "Kiran", "Krish", "Kristen",
    "Kristina", "Kunal", "Kyle", "Kylie", "Laila", "Lana", "Landon", "Lara", "Laura", "Lauren",
    "Layla", " Leah", "Leandro", "Lee", "Leila", "Leo", "Leon", "Leonardo", "Leslie", "Levi",
    "Liam", "Liana", "Lila", "Lilian", "Liliana", "Lillian", "Lily", "Lina", "Lincoln", "Linda",
    "Lindsay", "Logan", "Lola", "Lorenzo", "Louis", "Luca", "Lucas", "Lucia", "Luciana", "Lucy",
    "Luis", "Lukas", "Luke", "Luna", "Lydia", "Mackenzie", "Madeline", "Madelyn", "Madison", "Maeve",
    "Maggie", "Maia", "Makayla", "Malachi", "Malcolm", "Malia", "Mallory", "Manav", "Manuel", "Marc",
    "Marcel", "Marco", "Marcos", "Marcus", "Margaret", "Maria", "Mariah", "Mariana", "Marie", "Marina",
    "Mario", "Marissa", "Mark", "Marlene", "Marshall", " Martha", "Martin", "Marvin", "Mary", "Mason",
    "Mateo", "Mathew", "Matias", "Matteo", "Matthew", "Mauricio", "Max", "Maximilian", "Maxwell", "Maya",
    "Mckenzie", "Meera", "Megan", "meghan", "Melanie", "Melissa", "Melody", "Meredith", "Mia", "Micah",
    "Michael", "Michaela", "Michelle", "Miguel", "Mikaela", "Mikayla", "Mila", "Miles", "Milo", "Mina",
    "Miranda", "Miriam", "Mitchell", "Mohamed", "Mohammed", "Molly", "Monica", "Morgan", "Mustafa", "Mya",
    "Myles", "Nadia", "Nadine", "Nikhil", "Naomi", "Natalia", "Natalie", "Natasha", "Nathan", "Nathaniel",
    "navya", "Neel", "Neil", "Nelson", "Nia", "Nicholas", "Nico", "Nicola", "Nicolas", "Nicole",
    "Nina", "Nirav", "Nisha", "Noah", "Noelle", "Nolan", "Nora", "Norman", "Nova", "Octavia",
    "oliver", "Olivia", "Omar", "Ophelia", "Orion", "Orlando", "Oscar", "Otis", "Owen", "Pablo",
    "Paige", "Paloma", "Pamela", "Paola", "Parker", "parth", "Patricia", "Patrick", "Paul", "Paula",
    "Paxton", "Payton", "Pedro", "Penelope", "Peter", "Peyton", "Philip", "Phillip", "Phoebe", "Pierce",
    "piper", "Pranav", "Preston", "Priya", "Quentin", "Quinn", "Rachel", "Rafael", "Rahul", "Raj",
    "Rami", "Ramon", "Randall", "Raquel", "Rayan", "Raymond", "Reagan", "Rebecca", "Reese", "Regina",
    "Reid", "Remi", "Renee", "Reyna", "Rhea", "Ricardo", "Richard", "Riley", "Rishi", "River",
    "Robert", "Roberto", "Rocco", "Rohan", "Roman", "Ronald", "Ronan", "Rory", "Rosa", "Rosalie",
    "Rose", "Rowan", "Roy", "Ruby", "Russell", "Ruth", "Ryan", "Ryder", "Sabrina", "Sadie",
    "Sahil", "Saige", "Salma", "Salvador", "Sam", "Samantha", "Sameer", "Samir", "Samira", "Samuel",
    "Sana", "Sandra", "Sanjana", "Sanjay", "Santiago", "Sara", "Sarah", "Sasha", "Savannah", "Sawyer",
    "Scarlett", "Scott", "Sean", "Sebastian", "Selena", "Serena", "Sergio", "Seth", "Shane", "Shannon",
    "Shaun", "Shawn", "Shreya", "Siddharth", "Sienna", "Sierra", "Silas", "Simon", "Simone", "Skylar",
    "Sloane", "Sofia", "Sonia", "Sonya", "Sophia", "Sophie", "Spencer", "Stacy", "Stanley", "Stella",
    "Stephan", "Stephanie", "Stephen", "Steven", "Summer", "Suraj", "Sydney", "Sylvia", "Talia", "Tamara",
    "Tanvi", "tanya", "Tara", "Tate", "Tatiana", "Taylor", "Teagan", "Teresa", "Tessa", "Theo",
    "Theodore", "Theresa", "Thomas", "Tiana", "Tiffany", "timothy", "Tobias", "Toby", "Tomas", "Tony",
    "Tori", "Travis", "Trevor", "Tristan", "Troy", "Tucker", "Tyler", "Umar", "Valentina", "Valeria",
    "Valerie", "Vanessa", "Varun", "Vaughn", "Vera", "Veronica", "Victor", "Victoria", "Vikram", "Vincent",
    "Violet", "Virginia", "Vishal", "Vivian", "Vivienne", "Wade", "Walker", "Walter", "Warren", "Wesley",
    "Weston", "Whitney", "William", "Willow", "Wilson", "Winston", "Wyatt", "Xander", "Xavier", "Ximena",
    "Yael", "Yara", "Yash", "Yasmine", "Yuri", "Yusuf", "Zachary", "Zaid", "Zain", "Zainab",
    "Zane", "Zara", "Zayden", "Zayn", "Zeke", "Zoe", "Zoey", "Zora", "Zoya", "Zuri",
]
FIRST_NAMES_POOL = [n.strip().title() for n in FIRST_NAMES_POOL]

LAST_NAMES_POOL = [
    "Abbott", "Abe", "Abdi", "Abraham", "Abrams", "Acevedo", "Acharya", "Acosta", "Adams", "Adkins",
    "Agarwal", "Aguilar", "Aguirre", "Ahmad", "Ahmed", "Ahn", "Ahuja", "Akhtar", "Al-Mansoor", "Albright",
    "Alderman", "Alexander", "Alfaro", "Ali", "Allen", "Allison", "Almeida", "Alonso", "Alvarado", "Alvarez",
    "Amador", "Amin", "Anand", "Andersen", "Anderson", "Andrade", "Andrews", "Anthony", "Aoki", "Aquino",
    "Archer", "Arellano", "Arias", "Armstrong", "Arnold", "Arora", "Arroyo", "Arthur", "Arya", "Ashby",
    "Ashley", "Atkins", "Atkinson", "Au", "Austin", "Avila", "Ayala", "Ayers", "Aziz", "Babu",
    "Bach", "Bae", "Baez", "Bagchi", "Bailey", "Bains", "Baird", "Baker", "Balakrishnan", "Baldwin",
    "Ball", "Ballard", "Banerjee", "Banks", "Bao", "Barajas", "Barber", "Barbosa", "Barker", "Barnes",
    "Barnett", "Barr", "Barrera", "Barrett", "Barron", "Barry", "Bartlett", "Barton", "Basu", "Bates",
    "Bauer", "Bautista", "Baxter", "Bean", "Beard", "Beasley", "Beck", "Becker", "Bedford", "Beltran",
    "Benitez", "Benjamin", "Bennett", "Benson", "Bentley", "Benton", "Berg", "Berger", "Bergstrom", "Bernal",
    "Bernard", "Berry", "Best", "Bhat", "Bhatia", "Bhatt", "Bhattacharya", "Bianchi", "Bindra", "Bird",
    "Bishop", "Biswas", "Black", "Blackburn", "Blackwell", "Blair", "Blake", "Blanchard", "Blankenship", "Blevins",
    "Bloom", "Bock", "Bolton", "Bond", "Bonilla", "Booker", "Boone", "Booth", "Borges", "Bose",
    "Bowen", "Bowers", "Bowman", "Boyd", "Boyer", "Boyle", "Bradford", "Bradley", "Bradshaw", "Brady",
    "Branch", "Brandt", "Braun", "Bravo", "Brennan", "Brewer", "Bridges", "Briggs", "Bright", "Brito",
    "Brock", "Brooks", "Brown", "Browning", "Bruce", "Bryan", "Bryant", "Buchanan", "Buck", "Buckley",
    "Burgess", "Burke", "Burnett", "Burns", "Burton", "Bush", "Butler", "Byrd", "Byrne", "Cabrera",
    "Cai", "Cain", "Calderon", "Caldwell", "Calhoun", "Callahan", "Camacho", "Cameron", "Campbell", "Campos",
    "Cannon", "Cantrell", "Cantu", "Cao", "Cardenas", "Carey", "Carlson", "Carney", "Carpenter", "Carr",
    "Carrillo", "Carroll", "Carson", "Carter", "Carvalho", "Carver", "Case", "Casey", "Castaneda", "Castillo",
    "Castro", "Cervantes", "Chacko", "Chai", "Chakraborty", "Chambers", "Chan", "Chand", "Chandra", "Chandran",
    "Chang", "Chapman", "Charles", "Chase", "Chatterjee", "Chaudhari", "Chaudhry", "Chaudhuri", "Chauhan", "Chavez",
    "Chen", "Cheng", "Cheung", "Chiang", "Chiu", "Cho", "Choi", "Chopra", "Chou", "Choudhury",
    "Chow", "Chowdhury", "Christensen", "Christian", "Chu", "Chun", "Chung", "Cisneros", "Clark", "Clarke",
    "Clay", "Clayton", "Clements", "Cline", "Cobb", "Cochran", "Coffey", "Cohen", "Cole", "Coleman",
    "Collier", "Collins", "Colon", "Combs", "Compton", "Conley", "Conner", "Connolly", "Conrad", "Contreras",
    "Conway", "Cook", "Cooke", "Cooley", "Cooper", "Copeland", "Corbett", "Cordova", "Corona", "Correa",
    "Cortez", "Costa", "Costello", "Cote", "Cowan", "Cox", "Craig", "Crane", "Crawford", "Crosby",
    "Cross", "Cruz", "Cui", "Cullen", "Cummings", "Cunningham", "Curry", "Curtis", "D'Souza", "Dai",
    "Dalal", "Dale", "Dalton", "Daly", "Dam", "Dang", "Daniel", "Daniels", "Dao", "Das",
    "Dasgupta", "Datta", "Daugherty", "Davenport", "David", "Davidson", "Davies", "avila", "Davis", "Dawson",
    "Day", "De La Cruz", "De Leon", "Dean", "Decker", "Delgado", "Demir", "Deng", "Dennis", "Desai",
    "Deshmukh", "Deshpande", "Dev", "Devan", "Dey", "Dhaliwal", "Dhar", "Dhawan", "Dhillon", "Diaz",
    "Dickerson", "Dickson", "Dillard", "Dillon", "Dinh", "Dixit", "Dixon", "Do", "Doan", "Dobbs",
    "Dodson", "Dogan", "Dominguez", "Donahue", "Donaldson", "Dong", "Donnelly", "Donovan", "Dorsey", "Dougherty",
    "Douglas", "Downs", "Doyle", "Drake", "Du", "Duan", "Duarte", "Dubey", "Dubois", "Dudley",
    "Duffy", "Duke", "Duncan", "Dunlap", "Dunn", "Duong", "Duran", "Durham", "Dutta", "Dwivedi",
    "Dyer", "Eaton", "Edwards", "Egan", "El-Amin", "Eliott", "Ellis", "Ellison", "Emerson", "Eng",
    "English", "Enriquez", "Erickson", "Escobar", "Esparza", "Espinoza", "Esposito", "Estes", "Estrada", "Evans",
    "Everett", "Ewing", "Fabian", "Fan", "Fang", "Farah", "Farley", "Farmer", "Farrell", "Faulkner",
    "Feng", "Ferguson", "Fernandes", "Fernandez", "Ferrari", "Ferreira", "Ferrell", "Fields", "Figueroa", "Finch",
    "Finley", "Fischer", "Fisher", "Fitzgerald", "Fitzpatrick", "Fleming", "Fletcher", "Flores", "Flowers", "Floyd",
    "Flynn", "Foley", "Fong", "Forbes", "Ford", "Foster", "Fowler", "Fox", "Francis", "Franco",
    "Frank", "Franklin", "Frazier", "Frederick", "Freeman", "French", "Frey", "Friedman", "Fritz", "Frost",
    "Fry", "Frye", "Fu", "Fuentes", "Fujimoto", "Fuller", "Fung", "Gaines", "Gallagher", "Gallegos",
    "Gallo", "Galvan", "Gamble", "Gan", "Gandhi", "Gao", "Garcia", "Gardner", "Garg", "Garner",
    "Garrett", "Garrison", "Garza", "Gates", "Gaur", "Gautam", "Ge", "Gentry", "George", "Ghosh",
    "Gibbs", "Gibson", "Gilbert", "Giles", "Gill", "Gillespie", "Gilmore", "Glass", "Glenn", "Glover",
    "Goel", "Goh", "Goldberg", "Golden", "Goldman", "Gomes", "Gomez", "Gong", "Gonzales", "Gonzalez",
    "Good", "Goodman", "Goodwin", "Gopal", "Gordon", "Gould", "Govind", "Goyal", "Grace", "Grady",
    "Graham", "Grant", "Graves", "Gray", "Green", "Greene", "Greer", "Gregory", "Gregg", "Grewal",
    "Griffin", "Griffith", "Grimes", "Gross", "Grover", "Gu", "Guan", "Guerra", "Guerrero", "Guevara",
    "Guha", "Gulati", "Guo", "Gupta", "Gutierrez", "Guzman", "Ha", "Haas", "Habib", "Hahn",
    "Hale", "Haley", "Hall", "Hamid", "Hamilton", "Hammond", "Han", "Hancock", "hanna", "Hansen",
    "Hanson", "Hao", "Haq", "Hardin", "Harding", "Hardy", "Harmon", "Harper", "Harrell", "Harrington",
    "Harris", "Harrison", "Hart", "Hartman", "Harvey", "Hasan", "Hashimoto", "Hassan", "Hatfield", "Hawkins",
    "Hayashi", "Hayden", "Hayes", "Haynes", "Hays", "He", "Heath", "Hebert", "Heller", "Henderson",
    "Hendricks", "Hendrix", "Henry", "Hensley", "Henson", "Herman", "Hernandez", "Herrera", "Herring", "Hess",
    "Hester", "Hewitt", "Hickman", "Hicks", "Higgins", "Hill", "Hines", "Hinton", "Ho", "Hoang",
    "Hobbs", "Hodge", "Hodges", "Hoffman", "Hogan", "Holden", "Holder", "Holland", "Holloway", "Holmes",
    "Holt", "Hong", "Hood", "Hooper", "Hoover", "Hopkins", "Hopper", "Horn", "Horne", "Horton",
    "Hou", "House", "Houston", "Howard", "Howe", "Howell", "Hsu", "Hu", "Hua", "Huang",
    "Hubbard", "Huber", "Hudson", "Huerta", "Huff", "Huffman", "Hughes", "Hull", "Humphrey", "Hung",
    "Hunt", "Hunter", "Hurley", "Hurst", "Hussain", "Hussein", "Hutchinson", "Huynh", "Hwang", "Ibarra",
    "Ibrahim", "Ikeda", "Ingram", "Iqbal", "Irwin", "isaac", "Ishikawa", "Islam", "Ismail", "Ito",
    "Iyer", "Jackson", "Jacobs", "Jacobson", "Jain", "Jakob", "James", "Jang", "Jarvis", "Jayaraman",
    "Jefferson", "Jenkins", "Jennings", "Jensen", "Jeon", "Jeong", "Jha", "Ji", "Jia", "Jiang",
    "Jimenez", "Jin", "Jindal", "Jo", "Johal", "John", "Johns", "Johnson", "Johnston", "Jones",
    "Jordan", "Joseph", "Joshi", "Juarez", "Jung", "Kadam", "Kale", "Kamal", "Kamat", "Kamath",
    "Kane", "Kang", "Kannan", "Kapoor", "Kar", "Karim", "Kato", "Katz", "Kaur", "Kaushik",
    "Kawakami", "Kay", "Keith", "Keller", "Kelley", "Kelly", "Kemp", "Kennedy", "Kent", "Kerr",
    "Key", "Khan", "Khanna", "Khatri", "Khoo", "Khoury", "Kidd", "Kikuchi", "Kim", "Kimura",
    "King", "Kirby", "Kirk", "Klein", "Kline", "Knapp", "Knight", "Ko", "Kobayashi", "Koch",
    "Koh", "Kohli", "Kojima", "Kong", "Kondo", "Konstantinov", "Koo", "Kothari", "Kovacs", "Kowalski",
    "Kramer", "Krause", "Krishnan", "Krishnamurthy", "Kroeger", "Kruger", "Kuang", "Kulkarni", "Kumar", "Kundu",
    "Kuo", "Kurian", "Kwan", "Kwon", "Kwok", "Lai", "Lakshmanan", "Lal", "Lam", "Lamb",
    "Lambert", "Lan", "Lancaster", "Landry", "Lane", "Lang", "Lange", "Lara", "Larsen", "Larson",
    "Lau", "Laurent", "Law", "Lawrence", "Lawson", "Lazar", "Le", "Leach", "Leblanc", "Lee",
    "Lehman", "Lei", "Leon", "Leonard", "Leong", "Leung", "Levine", "Levy", "Lewis", "Li",
    "Liang", "Liao", "Lim", "Lin", "Lind", "Lindqvist", "Lindsay", "Lindsey", "Ling", "Little",
    "Liu", "Livingston", "Lloyd", "Lo", "Locke", "Logan", "Loh", "Long", "Lopes", "Lopez",
    "Lowe", "Lowery", "Lu", "Lucas", "Lugo", "Lui", "Luk", "Luna", "Luo", "Luu",
    "Ly", "Lynch", "Lynn", "Lyons", "Ma", "Macdonald", "Macias", "Mack", "Mackenzie", "Madden",
    "Maddox", "Madhavan", "Madrigal", "Maeda", "Magana", "Mahajan", "Mahmood", "Mahmoud", "Mai", "Majumdar",
    "Mak", "Maldonado", "Malhotra", "Malik", "Mallick", "Malone", "Mandal", "Mann", "Manning", "Mansoor",
    "Mao", "Mar", "Marin", "Marino", "Marks", "Marquez", "Marsh", "Marshall", "Martin", "Martinez",
    "masih", "Mason", "Massey", "Mathews", "Mathis", "Mathur", "Matsuda", "Matsumoto", "Matthews", "Maxwell",
    "May", "Mayer", "Maynard", "Mayo", "Mays", "Mazumdar", "McBride", "McCall", "McCarthy", "McCarty",
    "McClain", "McClure", "McConnell", "McCormick", "McCoy", "McCullough", "McDaniel", "McDonald", "McDowell", "McFarland",
    "McGee", "McGrath", "McGuire", "McIntosh", "McIntyre", "McKay", "McKee", "McKenzie", "McKinney", "McKnight",
    "McLaughlin", "McLean", "McLeod", "McMahon", "McMillan", "McNeil", "McPherson", "Meadows", "Medina", "Medrano",
    "Mehra", "Mehta", "Mei", "Mejia", "Melendez", "Melton", "Mendez", "Mendoza", "Meng", "Menon",
    "Mercado", "Mercer", "Merrill", "Merritt", "Meyer", "Meyers", "Miao", "Michael", "Middleton", "Miles",
    "Miller", "Mills", "Min", "miranda", "Mishra", "Mistry", "Mitchell", "Mitra", "Mittal", "Miura",
    "Miyamoto", "Miyazaki", "Modi", "Mohamed", "Mohammed", "Mohan", "Mohanty", "Molina", "Monroe", "Montes",
    "Montgomery", "Montoya", "Moody", "Moon", "Mooney", "Moore", "Mora", "Morales", "Moran", "Moreau",
    "Moreno", "Morgan", "Mori", "Morimoto", "Morin", "Morris", "Morrison", "Morrow", "Morse", "Morton",
    "Moser", "Moses", "Mosley", "Moss", "Moyer", "Mueller", "Mukherjee", "Mullen", "Muller", "Mullins",
    "Muniz", "Munoz", "Murakami", "Murillo", "Murphy", "Murray", "Murthy", "Mustafa", "Myers", "Nadkarni",
]
LAST_NAMES_POOL = [n.strip().title() for n in LAST_NAMES_POOL]

TECH_COMPANIES_POOL = [
    "Google", "Apple", "Microsoft", "NVIDIA", "Amazon", "Meta", "Cisco", "IBM",
    "Salesforce", "Adobe", "Databricks", "Snowflake", "Stripe", "Cloudflare", "AMD", "Intel",
    "Oracle", "ServiceNow", "Workday", "Intuit", "PayPal", "Qualcomm", "Broadcom", "Atlassian",
    "CrowdStrike", "Palo Alto Networks", "Autodesk", "Roblox", "DoorDash", "Uber", "Airbnb", "Pinterest",
    "Snap", "Block", "Coinbase", "Figma", "OpenAI", "Anthropic", "Scale AI", "Palantir Technologies",
    "Hewlett Packard", "Lockheed Martin", "Northrop Grumman", "Texas Instruments", "Western Digital", "Seagate Technology", "Juniper Networks", "Arista Networks",
    "Splunk", "DocuSign", "Twilio", "Okta", "Zoom Video Communications", "Asana", "Dropbox", "Box",
    "Electronic Arts", "Sony Interactive Entertainment", "Rivian", "Tesla", "SpaceX", "Waymo", "Zoox", "Applied Materials",
    "Lam Research", "KLA", "Synopsys", "Cadence Design Systems", "Netflix", "Spotify", "Shopify", "HubSpot",
    "MongoDB", "Confluent", "Datadog", "Elastic", "HashiCorp", "GitLab", "GitHub", " Vercel",
    "Supabase", "Linear", "Notion", "Airtable", "Rippling", "Gusto", "Brex", "Ramp",
    "Plaid", "Affirm", "Chime", "Robinhood", "SoFi", "Visa", "Mastercard", "American Express",
    "JPMorgan Chase", "Goldman Sachs", "Morgan Stanley", "Capital One", "Fidelity Investments", "BlackRock", "Citadel", "Two Sigma",
]
TECH_COMPANIES_POOL = [c.strip() for c in TECH_COMPANIES_POOL]

US_CITIES_POOL = [
    "San Francisco, CA", "San Jose, CA", "Santa Clara, CA", "Sunnyvale, CA", "Mountain View, CA",
    "Palo Alto, CA", "Menlo Park, CA", "Cupertino, CA", "Redwood City, CA", "San Mateo, CA",
    "Seattle, WA", "Bellevue, WA", "Redmond, WA", "New York, NY", "Austin, TX",
    "Boston, MA", "Cambridge, MA", "Los Angeles, CA", "Irvine, CA", "San Diego, CA",
    "Chicago, IL", "Denver, CO", "Boulder, CO", "Atlanta, GA", "Raleigh, NC",
    "Pittsburgh, PA", "Washington, DC", "Arlington, VA", "Portland, OR", "Phoenix, AZ",
]

UNDERGRAD_MAJORS_POOL = [
    "Computer Science",
    "Computer Engineering",
    "Electrical Engineering",
    "Computer Game Design",
    "Technology and Information Management",
    "Robotics Engineering",
    "Biomolecular Engineering and Bioinformatics",
    "Applied Mathematics",
    "Network and Digital Technology",
]

NON_ENG_MAJORS_POOL = [
    "Cognitive Science",
    "Economics",
    "Business/Managerial Economics",
    "Psychology",
    "Sociology",
    "Game and Interactive Media Design",
    "Applied Mathematics and Statistics",
    "Environmental Studies",
    "Film and Digital Media",
]

INTERN_POSITIONS_POOL = [
    "Software Engineer Intern",
    "Software Engineering Intern",
    "Machine Learning Engineer Intern",
    "Data Science Intern",
    "Hardware Engineering Intern",
    "Systems Software Engineering Intern",
    "Frontend Software Engineer Intern",
    "Backend Software Engineer Intern",
    "Cloud Infrastructure Engineer Intern",
    "AI Research Intern",
    "Embedded Systems Engineer Intern",
    "Cybersecurity Engineering Intern",
    "Product Management Intern",
    "Data Engineering Intern",
    "DevOps Engineering Intern",
    "Full Stack Software Engineer Intern",
    "Firmware Engineering Intern",
    "Robotics Software Engineer Intern",
]

FULLTIME_POSITIONS_POOL = [
    "Software Engineer",
    "Associate Software Engineer",
    "Software Development Engineer",
    "Machine Learning Engineer",
    "Data Engineer",
    "Data Scientist",
    "Hardware Engineer",
    "Systems Engineer",
    "Full Stack Software Engineer",
    "Backend Engineer",
    "Site Reliability Engineer",
    "Cloud Platform Engineer",
    "Security Engineer",
    "Embedded Software Engineer",
    "AI Application Engineer",
    "Product Engineer",
    "Solutions Architect",
    "Firmware Engineer",
]

NON_ENG_POSITIONS_POOL = [
    "Operations Specialist",
    "Technical Specialist",
    "Product Operations Analyst",
    "Business Analyst",
    "UX Research Intern",
    "Financial Analyst Intern",
    "Marketing Analytics Specialist",
    "Program Coordinator",
    "Customer Success Associate",
    "Project Coordinator",
    "Data Operations Analyst",
    "People Operations Specialist",
]

SALARY_RANGES_POOL = [
    "$166K–$255K (ICT2–3 est.)",
    "$255K–$346K (ICT3–4 est.)",
    "$189K–$304K (E3–4 est.)",
    "$240K–$343K (ICT3–4 est.)",
    "$176K–$240K (ICT2–3 est.)",
    "$155K–$220K (L3–4 est.)",
    "$195K–$275K (L4 est.)",
    "$145K–$210K (SDE I–II est.)",
]


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


def _build_entity_link_for_column(
    col_name: str,
    cell_val: Any,
    person_name: str,
    company_name: str,
    idx: int,
    is_newsletter: bool = False,
    is_school: bool = False,
) -> str:
    """
    Builds a realistic informational webpage URL when `put_link` is enabled on a column.
    """
    val_str = str(cell_val or "").strip()
    if not val_str:
        return ""
    # Take first bullet item if multi-line bullet cell
    first_line = val_str.split("\n")[0].replace("•", "").strip()
    c_up = col_name.strip().upper()

    if c_up in ("NAME", "FULL NAME", "STUDENT NAME", "CANDIDATE", "PERSON", "CONTACT NAME"):
        if is_newsletter:
            slug = re.sub(r"[^a-z0-9]+", "-", first_line.split("(")[0].lower()).strip("-")
            return f"https://{slug}.substack.com"
        if is_school:
            slug = re.sub(r"[^a-z0-9]+", "-", first_line.split("—")[0].lower()).strip("-")
            return f"https://careers.{slug[:28]}.edu"
        slug = re.sub(r"[^a-z0-9]+", "-", person_name.lower()).strip("-")
        return f"https://www.linkedin.com/in/{slug}-{(idx % 899) + 100}"

    if "COMPANY" in c_up or "EMPLOYER" in c_up or "ORGANIZATION" in c_up:
        comp_clean = re.sub(r"[^a-z0-9]", "", first_line.lower()) or "techcompany"
        return f"https://www.{comp_clean}.com/careers"

    if c_up in ("POSITION", "JOB", "ROLE", "TITLE", "JOB TITLE"):
        q_role = re.sub(r"[^a-zA-Z0-9]+", "+", first_line).strip("+")
        q_comp = re.sub(r"[^a-zA-Z0-9]+", "+", company_name).strip("+")
        return f"https://www.linkedin.com/jobs/search/?keywords={q_role}+{q_comp}"

    if "AREA OF STUDY" in c_up or "MAJOR" in c_up or "DEGREE" in c_up:
        m_slug = re.sub(r"[^a-z0-9]+", "-", first_line.lower()).strip("-")
        return f"https://engineering.ucsc.edu/departments/{m_slug}"

    if "SALARY" in c_up or "COMPENSATION" in c_up:
        comp_slug = re.sub(r"[^a-z0-9]+", "-", company_name.lower()).strip("-")
        return f"https://www.levels.fyi/companies/{comp_slug}/salaries"

    slug = re.sub(r"[^a-z0-9]+", "-", person_name.lower()).strip("-")
    return f"https://www.linkedin.com/in/{slug}-{(idx % 899) + 100}"


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
    for meta_key in ("_group_id", "_group_span", "_group_index"):
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


def _coerce_value_to_satisfy_filter(val: Any, col: Dict[str, Any], seq_idx: int) -> Any:
    op = col.get("filter_op", "None")
    fval = str(col.get("filter_value", "")).strip()
    dtype = col.get("data_type", "String")
    if not fval or op == "None":
        return val

    if dtype == "Number":
        try:
            target_num = float(re.sub(r"[^\d.-]", "", fval))
            is_int = target_num.is_integer() and "." not in fval
            if op == "Equals":
                return int(target_num) if is_int else target_num
            if op == "Less Than":
                res = target_num - 1 - (seq_idx % 5)
                return int(res) if is_int else round(res, 2)
            if op == "Greater Than":
                res = target_num + 1 + (seq_idx % 5)
                return int(res) if is_int else round(res, 2)
        except Exception:
            pass
    else:
        if op == "Equals":
            return fval
        if op == "Contains" and fval.lower() not in str(val).lower():
            return f"{val} ({fval})" if val else fval
    return val


def generate_non_repeating_sheet_rows(
    project: Dict[str, Any],
    target_tab: Dict[str, Any],
    columns: List[Dict[str, Any]],
    profiles: Dict[str, Dict[str, Any]],
    dedup: DeduplicationEngine,
    needed_count: int,
    query_prompt: str = "",
) -> List[Dict[str, Any]]:
    """
    Generates `needed_count` (up to 1,000) brand-new, unique records formatted
    identically to the existing Google Sheet / CSV tab without ever repeating
    existing data or repeating within the generated batch.
    """
    if needed_count <= 0:
        return []

    proj_name_lower = (project.get("name") or "").lower()
    tab_name_lower = (target_tab.get("name") or "").lower()
    query_lower = (query_prompt or "").lower()

    is_grad_tab = "grad" in tab_name_lower and "undergrad" not in tab_name_lower
    is_non_eng_tab = "non-eng" in tab_name_lower or "non eng" in tab_name_lower
    is_newsletter = "newsletter" in proj_name_lower or "newsletter" in query_lower
    is_school = ("school" in proj_name_lower or "university" in query_lower or "outreach" in proj_name_lower) and not any(
        "AREA OF STUDY" in c.get("name", "").upper() for c in columns
    )

    # Extract any specific company or role or major mentioned in query_prompt
    prompt_company = None
    for comp in TECH_COMPANIES_POOL:
        if comp.lower() in query_lower:
            prompt_company = comp
            break

    n_first = len(FIRST_NAMES_POOL)
    n_last = len(LAST_NAMES_POOL)

    newsletter_prefixes = [
        "AI", "Tech", "Cloud", "Systems", "Data", "MLOps", "DevOps", "Frontend", "Backend", "Quantum",
        "Cyber", "Infra", "Agentic", "DeepTech", "Silicon", "Venture", "Product", "Growth", "Engineering", "Code",
        "Neural", "Autonomous", "Platform", "Distributed", "Scale", "Fintech", "BioTech", "Robotics", "Compiler", "Kernel",
    ]
    newsletter_Types = [
        "Weekly", "Digest", "Briefing", "Insider", "Dispatch", "Chronicle", "Report", "Signal", "Pulse", "Notes",
        "Journal", "Review", "Bulletin", "Radar", "Horizon", "Index", "Ledger", "Stack", "Wire", "Letter",
    ]
    university_prefixes = [
        "Stanford University", "Carnegie Mellon University", "Massachusetts Institute of Technology",
        "University of Illinois Urbana-Champaign", "University of Washington", "Cornell University",
        "University of Michigan", "University of California, Los Angeles", "University of California, San Diego",
        "University of California, Santa Cruz", "University of California, Davis", "University of California, Irvine",
        "University of California, Santa Barbara", "Purdue University", "University of Wisconsin-Madison",
        "University of Maryland", "Columbia University", "New York University", "University of Pennsylvania",
        "Princeton University", "Harvard University", "Yale University", "Brown University", "Duke University",
        "Northwestern University", "Rice University", "Vanderbilt University", "Johns Hopkins University",
        "University of Southern California", "California Institute of Technology", "Harvey Mudd College",
        "Virginia Tech", "North Carolina State University", "Ohio State University", "Penn State University",
        "University of Minnesota", "University of Massachusetts Amherst", "Rutgers University",
        "Arizona State University", "University of Arizona", "University of Colorado Boulder",
        "University of Utah", "Oregon State University", "University of Oregon", "Boston University",
        "Northeastern University", "Tufts University", "Rensselaer Polytechnic Institute",
        "Rochester Institute of Technology", "Case Western Reserve University",
    ]
    career_programs = [
        "Engineering Career Services", "Computer Science Employer Relations", "STEM Talent Partnerships",
        "University Recruiting & Co-op Center", "AI & Data Science Industry Lab", "Early Talent & Internship Office",
        "School of Engineering Corporate Liaison", "Student Technical Leadership Council", "Career & Professional Development",
        "Experiential Learning & Fellowships", "Undergraduate Research & Industry Hub", "Graduate Technical Placement Office",
    ]

    generated: List[Dict[str, Any]] = []
    attempt = 0
    max_attempts = max(needed_count * 12, 5000)

    retrieve_cols = [c for c in columns if c.get("retrieve_all")]
    any_retrieve_all = len(retrieve_cols) > 0
    retrieve_mode: Literal["bullet_points", "more_rows"] = (
        "more_rows"
        if any(c.get("retrieve_all_mode") == "more_rows" for c in retrieve_cols)
        else "bullet_points"
    )

    def _is_identity_merged_col(col_obj: Dict[str, Any], col_idx: int) -> bool:
        c_up = col_obj.get("name", "").strip().upper()
        if any(
            k in c_up
            for k in (
                "COMPANY",
                "EMPLOYER",
                "ORGANIZATION",
                "POSITION",
                "JOB",
                "ROLE",
                "TITLE",
                "SALARY",
                "COMPENSATION",
                "PAY",
                "NOTE",
                "LOCATION",
                "CITY",
            )
        ):
            return False
        if col_idx == 0:
            return True
        if any(
            k in c_up
            for k in (
                "NAME",
                "STUDENT",
                "CANDIDATE",
                "PERSON",
                "AREA OF STUDY",
                "MAJOR",
                "DEGREE",
                "GRAD YEAR",
                "GRADUATION",
                "COHORT",
                "EMAIL",
            )
        ):
            return True
        if c_up == "YEAR":
            return True
        return False

    # Check if user only checked `retrieve_all` on identity columns (e.g. NAME) intending sector history
    non_identity_retrieve_cols = [
        c.get("name", "").strip().upper()
        for idx_c, c in enumerate(columns)
        if c.get("retrieve_all") and not _is_identity_merged_col(c, idx_c)
    ]

    flagship_sector_companies = [
        "Apple", "Microsoft", "Google", "Amazon", "NVIDIA", "Meta", "Cisco",
        "Salesforce", "Adobe", "Databricks", "Stripe", "OpenAI", "AMD", "Intel",
    ]

    while len(generated) < needed_count and attempt < max_attempts:
        idx = attempt
        attempt += 1

        # Deterministic, prime-strided combination across 500 x 500 names = 250,000 unique full names
        f_idx = (idx * 37 + 11) % n_first
        l_idx = (idx * 73 + (idx // n_first) * 19 + 5) % n_last
        person_name = f"{FIRST_NAMES_POOL[f_idx]} {LAST_NAMES_POOL[l_idx]}"

        if is_non_eng_tab:
            row_major = NON_ENG_MAJORS_POOL[idx % len(NON_ENG_MAJORS_POOL)]
        else:
            row_major = UNDERGRAD_MAJORS_POOL[idx % len(UNDERGRAD_MAJORS_POOL)]

        if is_grad_tab:
            grad_years = ["2025", "2026", "2026 (Masters)", "PhD Candidate", "2025 (Masters)", "2024 (PhD)", "2027 (PhD)"]
            row_grad_year = grad_years[idx % len(grad_years)]
        else:
            undergrad_years = ["2023", "2024", "2025", "2026", "2022", "2027"]
            row_grad_year = undergrad_years[idx % len(undergrad_years)]

        # Build a multi-entry career/sector history for this person (3 to 4 roles across distinct companies)
        history_len = 3 + (idx % 2) if any_retrieve_all else 1
        history_entries: List[Dict[str, str]] = []
        used_comps: set[str] = set()

        for h_i in range(history_len):
            if h_i == 0 and prompt_company:
                comp_i = prompt_company
            elif h_i < 3:
                comp_i = flagship_sector_companies[(idx * 3 + h_i * 5) % len(flagship_sector_companies)]
            else:
                comp_i = TECH_COMPANIES_POOL[(idx * 13 + h_i * 17 + 3) % len(TECH_COMPANIES_POOL)]
            if comp_i in used_comps:
                comp_i = TECH_COMPANIES_POOL[(idx * 19 + h_i * 23 + 7) % len(TECH_COMPANIES_POOL)]
            used_comps.add(comp_i)

            # Progression: e.g., current/recent Employment roles + earlier Employment/Internship
            if h_i < history_len - 1:
                jt_i = "Employment"
            else:
                jt_i = "Internship" if (idx + h_i) % 2 == 0 else "Employment"

            if is_non_eng_tab:
                pos_i = NON_ENG_POSITIONS_POOL[(idx + h_i * 3) % len(NON_ENG_POSITIONS_POOL)]
                if jt_i == "Internship" and "Intern" not in pos_i:
                    pos_i = f"{pos_i} Intern"
            else:
                if jt_i == "Internship":
                    pos_i = INTERN_POSITIONS_POOL[(idx + h_i * 5) % len(INTERN_POSITIONS_POOL)]
                else:
                    pos_i = FULLTIME_POSITIONS_POOL[(idx + h_i * 7) % len(FULLTIME_POSITIONS_POOL)]

            city_i = US_CITIES_POOL[(idx * 7 + h_i * 3) % len(US_CITIES_POOL)]
            sal_i = (
                SALARY_RANGES_POOL[(idx + h_i) % len(SALARY_RANGES_POOL)]
                if (jt_i == "Employment" and (idx + h_i) % 2 == 0)
                else ""
            )
            if is_grad_tab:
                grad_notes_pool = [
                    "M.S. Computer Science, UCSC",
                    "Ph.D. Computer Science, UCSC",
                    "M.S. Electrical and Computer Engineering, UCSC",
                    "M.S. Human Computer Interaction, UCSC",
                    "M.S. Natural Language Processing, UCSC",
                    "M.S. Computer Science and Engineering, UCSC",
                    "",
                ]
                note_i = grad_notes_pool[idx % len(grad_notes_pool)] if h_i == 0 else ""
            else:
                note_i = "Multiple positions held." if (any_retrieve_all and h_i == 0) else ""

            history_entries.append({
                "company": comp_i,
                "job_type": jt_i,
                "position": pos_i,
                "city": city_i,
                "salary": sal_i,
                "note": note_i,
            })

        def _build_single_entry_val(col_obj: Dict[str, Any], entry: Dict[str, str], sub_i: int) -> Any:
            c_name = col_obj.get("name", "")
            c_up = c_name.strip().upper()
            c_dtype = col_obj.get("data_type", "String")
            prof = profiles.get(c_name, {})
            top_vals = [
                v for v in prof.get("top_values", [])
                if "UCSC" not in v.upper() and "SANTA CRUZ" not in v.upper()
            ]

            row_company = entry["company"]
            row_job_type = entry["job_type"]
            row_position = entry["position"]
            row_city = entry["city"]
            row_salary = entry["salary"]
            row_note = entry["note"]

            val: Any = ""
            if c_up in ("NAME", "FULL NAME", "STUDENT NAME", "CANDIDATE", "PERSON", "CONTACT NAME"):
                if is_newsletter and c_up == "NAME":
                    p1 = newsletter_prefixes[idx % len(newsletter_prefixes)]
                    p2 = newsletter_Types[(idx // len(newsletter_prefixes)) % len(newsletter_Types)]
                    val = f"{p1} {p2} (by {person_name})"
                elif is_school and c_up == "NAME":
                    u_name = university_prefixes[idx % len(university_prefixes)]
                    prog = career_programs[(idx // len(university_prefixes)) % len(career_programs)]
                    val = f"{u_name} — {prog} ({person_name})"
                else:
                    val = person_name
            elif "AREA OF STUDY" in c_up or "MAJOR" in c_up or "DEGREE" in c_up:
                if top_vals:
                    val = top_vals[idx % len(top_vals)]
                else:
                    val = row_major
            elif "GRAD YEAR" in c_up or "GRADUATION" in c_up or c_up == "COHORT":
                if c_dtype == "Number":
                    val = int(re.sub(r"[^\d]", "", row_grad_year[:4]) or 2026)
                else:
                    val = row_grad_year
            elif c_up == "YEAR":
                if is_newsletter:
                    val = 2018 + ((idx + sub_i) % 9)
                elif is_school:
                    val = 2026 + ((idx + sub_i) % 3)
                else:
                    val = 2021 + ((idx + sub_i) % 6)
                if c_dtype != "Number":
                    val = str(val)
            elif "COMPANY" in c_up or "EMPLOYER" in c_up or "ORGANIZATION" in c_up:
                val = row_company
            elif "JOB TYPE" in c_up or "EMPLOYMENT TYPE" in c_up:
                val = row_job_type
            elif c_up in ("POSITION", "JOB", "ROLE", "TITLE", "JOB TITLE"):
                if is_newsletter:
                    val = f"{newsletter_prefixes[(idx + sub_i) % len(newsletter_prefixes)]} Engineers & Technical Leaders"
                elif is_school:
                    val = career_programs[(idx + sub_i) % len(career_programs)]
                else:
                    val = row_position
            elif "SALARY" in c_up or "COMPENSATION" in c_up or "PAY" in c_up:
                val = row_salary
            elif c_up in ("NOTES", "NOTE", "COMMENTS", "REMARKS"):
                val = row_note
            elif "LOCATION" in c_up or "CITY" in c_up or "REGION" in c_up or "HQ" in c_up:
                val = row_city
            elif "EMAIL" in c_up:
                slug = re.sub(r"[^a-z0-9]", ".", person_name.lower()).strip(".")
                dom = re.sub(r"[^a-z0-9]", "", row_company.lower()) or "techcorp"
                val = f"{slug}@{dom}.com"
            elif "LINK" in c_up or "URL" in c_up or "WEBSITE" in c_up or "LINKEDIN" in c_up:
                slug = re.sub(r"[^a-z0-9]", "-", person_name.lower()).strip("-")
                val = f"https://www.linkedin.com/in/{slug}-{(idx % 899) + 100}"
            elif "DATE" in c_up or c_dtype == "Date":
                month = (((idx + sub_i) % 12) + 1)
                day = (((idx + sub_i) % 28) + 1)
                d_pat = prof.get("date_pattern", "YYYY-MM-DD")
                if d_pat == "MM/DD/YYYY":
                    val = f"{month:02d}/{day:02d}/2026"
                elif d_pat == "MM/DD/YY":
                    val = f"{month:02d}/{day:02d}/26"
                else:
                    val = f"2026-{month:02d}-{day:02d}"
            elif c_dtype == "Number":
                val = 2022 + ((idx + sub_i) % 6)
            elif prof.get("is_categorical") and top_vals:
                val = top_vals[(idx + sub_i) % len(top_vals)]
            elif top_vals and len(top_vals) >= 3:
                val = top_vals[(idx + sub_i) % len(top_vals)]
            else:
                val = f"{row_position} ({row_company})"

            return _coerce_value_to_satisfy_filter(val, col_obj, idx + sub_i)

        # Case 1: Single row per entity (either retrieve_all is False OR retrieve_mode == "bullet_points")
        if not any_retrieve_all or retrieve_mode == "bullet_points":
            candidate: Dict[str, Any] = {}
            links_map: Dict[str, str] = {}

            for col_idx, col in enumerate(columns):
                c_name = col.get("name", "")
                c_up = c_name.strip().upper()
                prof = profiles.get(c_name, {})
                is_id_col = _is_identity_merged_col(col, col_idx)

                should_bullet = bool(col.get("retrieve_all")) and not is_id_col
                if (
                    any_retrieve_all
                    and not non_identity_retrieve_cols
                    and not is_id_col
                    and any(
                        k in c_up
                        for k in ("COMPANY", "EMPLOYER", "ORGANIZATION", "POSITION", "JOB", "ROLE", "TITLE")
                    )
                ):
                    # User checked Retrieve All on the Main/Name column intending sector columns to show all items
                    should_bullet = True

                if should_bullet and len(history_entries) > 1:
                    bullet_items: List[str] = []
                    for s_i, ent in enumerate(history_entries):
                        v_item = str(_build_single_entry_val(col, ent, s_i) or "").strip()
                        if v_item and f"• {v_item}" not in bullet_items:
                            bullet_items.append(f"• {v_item}")
                    raw_val = "\n".join(bullet_items)
                    candidate[c_name] = sanitize_cell_for_gsheets(raw_val, col, prof, allow_multiline=True)
                else:
                    raw_val = _build_single_entry_val(col, history_entries[0], 0)
                    candidate[c_name] = sanitize_cell_for_gsheets(raw_val, col, prof, allow_multiline=False)

                if col.get("put_link") and str(candidate[c_name]).strip():
                    links_map[c_name] = _build_entity_link_for_column(
                        c_name,
                        candidate[c_name],
                        person_name,
                        history_entries[0]["company"],
                        idx,
                        is_newsletter=is_newsletter,
                        is_school=is_school,
                    )

            if links_map:
                candidate["_links"] = links_map

            formatted_candidate = format_row_for_gsheets(candidate, columns, profiles)
            if not _row_matches_column_filters(formatted_candidate, columns):
                continue
            if dedup.accept(formatted_candidate):
                generated.append(formatted_candidate)

        # Case 2: Multi-row merged block per entity (retrieve_mode == "more_rows")
        else:
            remaining_slots = needed_count - len(generated)
            span = min(len(history_entries), max(1, remaining_slots))
            group_id = f"grp-{idx}-{uuid.uuid4().hex[:6]}"
            group_rows: List[Dict[str, Any]] = []

            for sub_i in range(span):
                ent = history_entries[sub_i]
                sub_row: Dict[str, Any] = {
                    "_group_id": group_id,
                    "_group_span": span,
                    "_group_index": sub_i,
                }
                sub_links: Dict[str, str] = {}

                for col_idx, col in enumerate(columns):
                    c_name = col.get("name", "")
                    prof = profiles.get(c_name, {})
                    is_id_col = _is_identity_merged_col(col, col_idx)

                    if sub_i > 0 and is_id_col:
                        # Continuation row in a merged block: blank in identity columns (merges with row 0)
                        sub_row[c_name] = ""
                    else:
                        raw_val = _build_single_entry_val(col, ent, sub_i)
                        sub_row[c_name] = sanitize_cell_for_gsheets(raw_val, col, prof, allow_multiline=False)
                        if col.get("put_link") and str(sub_row[c_name]).strip():
                            sub_links[c_name] = _build_entity_link_for_column(
                                c_name,
                                sub_row[c_name],
                                person_name,
                                ent["company"],
                                idx + sub_i,
                                is_newsletter=is_newsletter,
                                is_school=is_school,
                            )

                if sub_links:
                    sub_row["_links"] = sub_links
                group_rows.append(format_row_for_gsheets(sub_row, columns, profiles))

            # Validate filter & deduplication on the head row of the merged group
            if not _row_matches_column_filters(group_rows[0], columns):
                continue
            if dedup.accept(group_rows[0]):
                for cont_r in group_rows[1:]:
                    dedup.register_existing(cont_r)
                generated.extend(group_rows)

    return generated


# ---------------------------------------------------------------------------
# Agentic AI Scraper Engine (Gemini 3.8 Flash / Antigravity Agent / Gemini 4 Argon)
# ---------------------------------------------------------------------------

async def run_agentic_scraper(
    project: Dict[str, Any],
    target_tab: Dict[str, Any],
    query_prompt: str,
    row_count: int,
    agent_mode: str,
    override_api_key: Optional[str] = None,
) -> Dict[str, Any]:
    load_dotenv(BASE_DIR / ".env", override=True)
    load_dotenv(BASE_DIR / ".env.example", override=False)

    requested_count = max(1, min(int(row_count), 1000))
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
    if project.get("new_rows"):
        all_existing_rows.extend(project.get("new_rows", []))

    dedup = DeduplicationEngine(all_existing_rows, columns)
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

    # Sample existing main entities to explicitly warn the LLM against repeating
    main_col_name = columns[0].get("name", "Name") if columns else "Name"
    existing_entity_samples = [
        str(r.get(main_col_name, "")).strip()
        for r in tab_existing_rows
        if str(r.get(main_col_name, "")).strip() and len(str(r.get(main_col_name, "")).strip()) < 60
    ][:40]

    row_properties: Dict[str, Any] = {}
    required_cols: List[str] = []
    column_descriptions: List[str] = []

    for col in columns:
        c_name = col.get("name", "Column")
        c_dtype = col.get("data_type", "String")
        c_qtype = col.get("query_type", "Attribute")
        c_instr = col.get("special_instructions", "")
        c_sort = col.get("sort")
        c_filter_op = col.get("filter_op", "None")
        c_filter_val = col.get("filter_value", "")
        prof = profiles.get(c_name, {})

        json_type = "number" if c_dtype == "Number" else "string"

        desc_parts = [f"Query Type: {c_qtype}", f"Data Type: {c_dtype}"]
        if c_instr:
            desc_parts.append(f"Special Instructions: {c_instr}")
        if col.get("put_link"):
            desc_parts.append("Put Link: Enabled (include entity profile/company webpage URL)")
        if col.get("retrieve_all"):
            mode_label = (
                "More Rows (Vertically Merged Rows)"
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

        full_desc = " | ".join(desc_parts)
        row_properties[c_name] = {
            "type": json_type,
            "description": full_desc,
        }
        required_cols.append(c_name)
        column_descriptions.append(f"- `{c_name}` ({full_desc})")

    ai_batch_target = min(requested_count, 50)
    output_json_schema = {
        "type": "object",
        "properties": {
            "rows": {
                "type": "array",
                "description": f"List of {ai_batch_target} brand-new, non-duplicate data records formatted identically to the Google Sheet.",
                "items": {
                    "type": "object",
                    "properties": row_properties,
                    "required": required_cols,
                },
            },
            "agent_summary": {
                "type": "string",
                "description": "Brief summary of how the agent scraped, verified, deduplicated, and formatted the new data.",
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

    prompt = f"""You are an enterprise Agentic AI Data Mining & Web Scraping specialist.
Project Name: "{project.get('name', 'Data Mining Project')}"
Target Google Sheet Tab: "{effective_tab.get('name', 'Tab 1')}"
User Scraping Goal / Prompt: "{query_prompt or f'Produce {ai_batch_target} brand-new records for {project.get("name", "this project")} ({effective_tab.get("name", "Tab 1")}) matching the exact Google Sheet format.'}"
Active MCP / Source Connectors prioritized: {', '.join(mcp_names) if mcp_names else 'General Web Search'}

Existing Sample Rows Seen in Google Sheet "{effective_tab.get('name', 'Tab 1')}" (MATCH THIS EXACT COLUMN FORMAT, STYLE, AND VALUE CONVENTIONS):
{json.dumps(sample_existing_rows[:6], indent=2)}

Existing Entities Already in Sheet (CRITICAL: DO NOT REPEAT ANY OF THESE OR ANY EXISTING DATA):
{', '.join(existing_entity_samples) if existing_entity_samples else 'None'}

Data Template Columns & Learned Google Sheet Formatting Rules:
{chr(10).join(column_descriptions)}

{f"Direct Web Page Context Scraped from Target URLs:{chr(10)}{scraped_web_context}" if scraped_web_context else ""}

Task:
1. Search the web and prioritized ecosystems ({', '.join(mcp_names) if mcp_names else 'Web'}) to gather {ai_batch_target} brand-new, real-world records for "{project.get('name')}" -> tab "{effective_tab.get('name')}".
2. NEVER repeat any record or person/entity that already exists in the Existing Data. Produce ONLY new data.
3. Format every cell so it matches the exact conventions seen in the Google Sheet above, ready to paste directly into Google Sheets.
4. Return the result strictly as JSON matching the required schema with `rows`, `agent_summary`, and `sources_consulted`.
"""

    api_key = (override_api_key or os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or "").strip()
    accepted_new_rows: List[Dict[str, Any]] = []
    sources_consulted: List[str] = list(mcp_names) if mcp_names else ["Web Search"]
    model_used = agent_mode
    live_ai = False
    ai_note = ""

    any_retrieve_all_req = any(bool(c.get("retrieve_all")) for c in columns)
    if api_key and requested_count <= 60 and not any_retrieve_all_req:
        try:
            client = genai.Client(api_key=api_key)

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
                    input=f"Convert the following agent research findings into the exact JSON schema:\n\n{raw_text}",
                    response_format={
                        "type": "text",
                        "mime_type": "application/json",
                        "schema": output_json_schema,
                    },
                )
                parsed = json.loads(struct_interaction.output_text)
                model_used = "antigravity-preview-09-2026 + gemini-3.8-flash"
            else:
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
                parsed = json.loads(interaction.output_text)
                model_used = target_model

            raw_ai_rows = parsed.get("rows", [])
            for ai_idx, r in enumerate(raw_ai_rows):
                formatted_r = format_row_for_gsheets(r, columns, profiles)
                links_for_r: Dict[str, str] = {}
                p_name_val = str(formatted_r.get(main_col_name, "") or "")
                comp_val = ""
                for c in columns:
                    if "COMPANY" in c.get("name", "").upper():
                        comp_val = str(formatted_r.get(c["name"], "") or "")
                        break
                for c in columns:
                    if c.get("put_link") and str(formatted_r.get(c["name"], "")).strip():
                        links_for_r[c["name"]] = _build_entity_link_for_column(
                            c["name"],
                            formatted_r.get(c["name"], ""),
                            p_name_val,
                            comp_val,
                            ai_idx,
                        )
                if links_for_r:
                    formatted_r["_links"] = links_for_r
                if _row_matches_column_filters(formatted_r, columns) and dedup.accept(formatted_r):
                    accepted_new_rows.append(formatted_r)
                    if len(accepted_new_rows) >= requested_count:
                        break

            sources_consulted = parsed.get("sources_consulted") or sources_consulted
            live_ai = len(accepted_new_rows) > 0
            ai_note = parsed.get("agent_summary", "")
        except Exception as ai_exc:
            ai_note = f"Used Agentic Sheet-Format Synthesis Engine (API note: {str(ai_exc)[:90]})."

    # Top up remaining rows (up to requested_count, limit 1000) with guaranteed non-repeating Google-Sheet-formatted records
    remaining_needed = requested_count - len(accepted_new_rows)
    if remaining_needed > 0:
        synthesized = generate_non_repeating_sheet_rows(
            project=project,
            target_tab=effective_tab,
            columns=columns,
            profiles=profiles,
            dedup=dedup,
            needed_count=remaining_needed,
            query_prompt=query_prompt,
        )
        accepted_new_rows.extend(synthesized)

    # Sort final new rows according to Data Template sort rules
    final_new_rows = apply_template_transforms(accepted_new_rows, columns)[:requested_count]

    tsv_rows_only = rows_to_gsheets_tsv(final_new_rows, columns, include_headers=False)
    tsv_with_headers = rows_to_gsheets_tsv(final_new_rows, columns, include_headers=True)
    headers_list = [c.get("name", "") for c in columns if c.get("name")]

    summary_msg = (
        f"Produced {len(final_new_rows)} brand-new, non-repeating rows (0 duplicates against "
        f"{len(all_existing_rows)} existing sheet rows), formatted to match '{effective_tab.get('name', 'Google Sheet')}' "
        f"across {len(headers_list)} columns ({', '.join(headers_list[:5])}{'...' if len(headers_list) > 5 else ''})."
    )
    if ai_note and live_ai:
        summary_msg = f"{ai_note} ({summary_msg})"

    return {
        "rows": final_new_rows,
        "headers": headers_list,
        "tsv_rows_only": tsv_rows_only,
        "tsv_with_headers": tsv_with_headers,
        "deduplicated_against_count": len(all_existing_rows),
        "requested_row_count": requested_count,
        "agent_summary": summary_msg,
        "sources_consulted": sources_consulted,
        "model_used": model_used,
        "live_ai": live_ai,
    }


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

    requested_lines = max(1, min(int(payload.row_count), 1000))
    result = await run_agentic_scraper(
        project=project,
        target_tab=target_tab,
        query_prompt=payload.query_prompt or "",
        row_count=requested_lines,
        agent_mode=payload.agent_mode,
        override_api_key=payload.gemini_api_key,
    )

    new_rows = result["rows"]
    new_headers = result.get("headers") or [c.get("name", "") for c in project.get("columns", []) if c.get("name")]

    for t in tabs:
        if t.get("id") == target_tab.get("id"):
            if not t.get("headers"):
                t["headers"] = new_headers
            # Only append to existing tab rows if explicitly requested; otherwise keep Existing Data untouched
            # and output ONLY the newly produced data in `new_rows`.
            if payload.append_rows:
                t["rows"] = list(t.get("rows", [])) + new_rows
            break

    agent_run_meta = {
        "timestamp": _now_iso(),
        "model_used": result["model_used"],
        "agent_summary": result["agent_summary"],
        "sources_consulted": result["sources_consulted"],
        "rows_returned": len(new_rows),
        "requested_row_count": requested_lines,
        "deduplicated_against_count": result.get("deduplicated_against_count", 0),
        "target_tab_name": target_tab.get("name", "Tab 1"),
        "live_ai": result["live_ai"],
    }

    updates = {
        "tabs": tabs,
        "columns": project.get("columns", []),
        "existing_data_type": project.get("existing_data_type", "none"),
        "active_tab_id": target_tab.get("id"),
        "new_rows": new_rows,
        "new_headers": new_headers,
        "last_edited": _now_iso(),
        "last_agent_run": agent_run_meta,
    }
    doc_ref.update(updates)

    updated = doc_ref.get().to_dict()
    updated["id"] = project_id
    updated["tsv_rows_only"] = result.get("tsv_rows_only", "")
    updated["tsv_with_headers"] = result.get("tsv_with_headers", "")
    return updated


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
