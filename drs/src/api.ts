import { auth } from './firebase'

export type SortDirection = 'Ascending' | 'Descending' | null
export type QueryType = 'Main' | 'Attribute'
export type DataType = 'String' | 'Number' | 'Date'
export type FilterOp = 'Less Than' | 'Greater Than' | 'Equals' | 'Contains' | 'None'
export type RetrieveAllMode = 'bullet_points' | 'more_rows'

export type SheetRowValue = string | number | Record<string, string> | undefined
export type SheetRow = Record<string, SheetRowValue>

export interface ColumnTemplate {
  id: string
  name: string
  special_instructions: string
  sort: SortDirection
  query_type: QueryType
  data_type: DataType
  filter_op: FilterOp
  filter_value: string
  put_link?: boolean
  retrieve_all?: boolean
  retrieve_all_mode?: RetrieveAllMode
}

export interface SheetTab {
  id: string
  name: string
  gid?: string | null
  headers?: string[]
  rows: SheetRow[]
}

export interface MCPConnector {
  id: string
  name: string
  icon: 'linkedin' | 'handshake' | 'custom' | string
  enabled: boolean
  server_url?: string
  description?: string
}

export interface AgentRunMeta {
  timestamp: string
  model_used: string
  agent_summary: string
  sources_consulted: string[]
  rows_returned: number
  requested_row_count?: number
  deduplicated_against_count?: number
  target_tab_name?: string
  live_ai: boolean
}

export interface GSheetSyncMeta {
  synced: boolean
  mode?: string
  spreadsheet_title?: string
  tab_count?: number
  worksheet?: string
  rows_written?: number
  can_write?: boolean
  reason?: string
}

export interface Project {
  id: string
  name: string
  owner_uid: string
  owner_email?: string
  created_at?: string
  last_edited: string
  existing_data_source: string
  existing_data_type: 'csv' | 'gsheets' | 'none'
  active_tab_id: string
  tabs: SheetTab[]
  columns: ColumnTemplate[]
  mcps: MCPConnector[]
  new_rows?: SheetRow[]
  new_headers?: string[]
  tsv_rows_only?: string
  tsv_with_headers?: string
  last_agent_run?: AgentRunMeta | null
  gsheet_sync?: GSheetSyncMeta | null
}

export function getApiBase(): string {
  const stored =
    typeof window !== 'undefined' ? localStorage.getItem('DRS_API_BASE_URL') || '' : ''
  const raw = (stored || import.meta.env.VITE_API_BASE_URL || '').trim().replace(/\/+$/, '')
  // Strip trailing /api so that appending /api/... never results in /api/api/...
  return raw.replace(/\/api$/, '')
}

async function getAuthHeaders(includeJsonContentType = true): Promise<Record<string, string>> {
  const currentUser = auth.currentUser
  if (!currentUser) {
    throw new Error('User is not authenticated')
  }
  const token = await currentUser.getIdToken()
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
  }
  if (includeJsonContentType) {
    headers['Content-Type'] = 'application/json'
  }
  return headers
}

export async function fetchProjects(): Promise<Project[]> {
  const headers = await getAuthHeaders()
  const res = await fetch(`${getApiBase()}/api/projects`, { headers })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.detail || 'Failed to fetch projects')
  }
  const data = await res.json()
  return data.projects
}

export async function createProject(name: string): Promise<Project> {
  const headers = await getAuthHeaders()
  const res = await fetch(`${getApiBase()}/api/projects`, {
    method: 'POST',
    headers,
    body: JSON.stringify({ name }),
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.detail || 'Failed to create project')
  }
  return res.json()
}

export async function updateProject(
  projectId: string,
  updates: Partial<
    Pick<
      Project,
      'name' | 'existing_data_source' | 'existing_data_type' | 'active_tab_id' | 'tabs' | 'columns' | 'mcps' | 'new_rows'
    >
  > & { sync_to_gsheet?: boolean }
): Promise<Project> {
  const headers = await getAuthHeaders()
  const res = await fetch(`${getApiBase()}/api/projects/${projectId}`, {
    method: 'PUT',
    headers,
    body: JSON.stringify(updates),
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.detail || 'Failed to save project')
  }
  return res.json()
}

export async function updateSheetCell(
  projectId: string,
  tabId: string,
  rowIndex: number,
  columnName: string,
  value: string | number,
  syncToGsheet = true
): Promise<{ ok: boolean; gsheet_sync?: GSheetSyncMeta }> {
  const headers = await getAuthHeaders()
  const res = await fetch(`${getApiBase()}/api/projects/${projectId}/cell`, {
    method: 'POST',
    headers,
    body: JSON.stringify({
      tab_id: tabId,
      row_index: rowIndex,
      column_name: columnName,
      value,
      sync_to_gsheet: syncToGsheet,
    }),
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.detail || 'Failed to update cell')
  }
  return res.json()
}

export async function addRowToTab(
  projectId: string,
  tabId: string,
  row?: SheetRow,
  syncToGsheet = true
): Promise<Project> {
  const headers = await getAuthHeaders()
  const res = await fetch(`${getApiBase()}/api/projects/${projectId}/add-row`, {
    method: 'POST',
    headers,
    body: JSON.stringify({
      tab_id: tabId,
      row,
      sync_to_gsheet: syncToGsheet,
    }),
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.detail || 'Failed to add row')
  }
  return res.json()
}

export async function deleteProject(projectId: string): Promise<void> {
  const headers = await getAuthHeaders()
  const res = await fetch(`${getApiBase()}/api/projects/${projectId}`, {
    method: 'DELETE',
    headers,
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.detail || 'Failed to delete project')
  }
}

export async function uploadCsvFile(
  projectId: string,
  file: File,
  tabId: string,
  syncColumns = true
): Promise<Project> {
  const headers = await getAuthHeaders(false)
  const formData = new FormData()
  formData.append('file', file)
  formData.append('tab_id', tabId)
  formData.append('sync_columns', String(syncColumns))

  const res = await fetch(`${getApiBase()}/api/projects/${projectId}/upload-csv`, {
    method: 'POST',
    headers,
    body: formData,
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.detail || 'Failed to upload CSV file')
  }
  return res.json()
}

export async function importGoogleSheet(
  projectId: string,
  url: string,
  tabId?: string,
  syncColumns = true
): Promise<Project> {
  const headers = await getAuthHeaders()
  const res = await fetch(`${getApiBase()}/api/projects/${projectId}/import-gsheet`, {
    method: 'POST',
    headers,
    body: JSON.stringify({ url, tab_id: tabId, sync_columns: syncColumns }),
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.detail || 'Failed to import Google Sheet')
  }
  return res.json()
}

export async function syncGoogleSheet(
  projectId: string,
  direction: 'pull' | 'push',
  tabId?: string
): Promise<Project> {
  const headers = await getAuthHeaders()
  const res = await fetch(`${getApiBase()}/api/projects/${projectId}/sync-gsheet`, {
    method: 'POST',
    headers,
    body: JSON.stringify({ direction, tab_id: tabId }),
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.detail || `Failed to ${direction} Google Sheet`)
  }
  return res.json()
}

export async function runFetchDataAgent(
  projectId: string,
  options: {
    query_prompt?: string
    tab_id?: string
    row_count?: number
    append_rows?: boolean
    agent_mode?: 'gemini-3.8-flash' | 'antigravity-agent' | 'gemini-4-argon'
    gemini_api_key?: string
  }
): Promise<Project> {
  const headers = await getAuthHeaders()
  const res = await fetch(`${getApiBase()}/api/projects/${projectId}/fetch-data`, {
    method: 'POST',
    headers,
    body: JSON.stringify(options),
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.detail || 'Failed to run AI agent data fetch')
  }
  return res.json()
}
