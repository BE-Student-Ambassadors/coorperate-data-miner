import React, {
  useCallback,
  useDeferredValue,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react'
import {
  createUserWithEmailAndPassword,
  onAuthStateChanged,
  sendPasswordResetEmail,
  signInWithEmailAndPassword,
  signOut,
  type User,
} from 'firebase/auth'
import { auth } from './firebase'
import {
  addRowToTab,
  clearNewData,
  createProject,
  deleteProject,
  fetchProjects,
  getFetchProgress,
  importGoogleSheet,
  runFetchDataAgent,
  syncGoogleSheet,
  updateProject,
  updateSheetCell,
  uploadCsvFile,
  type ColumnTemplate,
  type DataType,
  type FetchProgress,
  type FilterOp,
  type MCPConnector,
  type Project,
  type QueryType,
  type RetrieveAllMode,
  type SheetRow,
  type SortDirection,
} from './api'
import './App.css'

const PAGE_SIZE = 100

function formatLastEdited(isoString?: string): string {
  if (!isoString) return ''
  try {
    const d = new Date(isoString)
    if (Number.isNaN(d.getTime())) return isoString
    return d.toLocaleString(undefined, {
      month: 'short',
      day: 'numeric',
      year: 'numeric',
      hour: 'numeric',
      minute: '2-digit',
    })
  } catch {
    return isoString
  }
}

function getCellValue(row: SheetRow, colName: string): string | number {
  if (!row || colName.startsWith('_')) return ''
  const direct = row[colName]
  if (typeof direct === 'string' || typeof direct === 'number') {
    return direct
  }
  const targetUpper = colName.trim().toUpperCase()
  for (const k of Object.keys(row)) {
    if (k.startsWith('_')) continue
    if (k.trim().toUpperCase() === targetUpper) {
      const v = row[k]
      if (typeof v === 'string' || typeof v === 'number') return v
    }
  }
  return ''
}

function getCellLink(
  row: SheetRow,
  colName: string,
  colTemplate?: ColumnTemplate
): string {
  if (!row) return ''
  const rawVal = String(getCellValue(row, colName) ?? '').trim()
  if (!rawVal) return ''
  // Ignore long instructional banner cells in row 0
  if (rawVal.length > 90 && rawVal.includes('Ex.')) return ''

  const linksObj = row._links
  if (linksObj && typeof linksObj === 'object') {
    if (typeof linksObj[colName] === 'string' && linksObj[colName].trim()) {
      return linksObj[colName].trim()
    }
    const targetUpper = colName.trim().toUpperCase()
    for (const [k, v] of Object.entries(linksObj)) {
      if (k.trim().toUpperCase() === targetUpper && typeof v === 'string' && v.trim()) {
        return v.trim()
      }
    }
  }

  // If the cell text itself is already a URL
  if (/^https?:\/\/\S+$/i.test(rawVal)) {
    return rawVal
  }

  // If `put_link` is enabled on this column template, synthesize the entity's profile/company link
  if (colTemplate?.put_link) {
    const firstLine = rawVal.split('\n')[0].replace(/^•\s*/, '').trim()
    if (!firstLine) return ''
    const cUp = colName.trim().toUpperCase()
    const slug = firstLine
      .toLowerCase()
      .replace(/\([^)]*\)/g, '')
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '')
    if (!slug) return ''

    if (
      cUp.includes('COMPANY') ||
      cUp.includes('EMPLOYER') ||
      cUp.includes('ORGANIZATION')
    ) {
      const domain = firstLine.toLowerCase().replace(/[^a-z0-9]/g, '') || 'company'
      return `https://www.${domain}.com/careers`
    }
    if (cUp.includes('POSITION') || cUp === 'JOB' || cUp.includes('ROLE') || cUp.includes('TITLE')) {
      const q = encodeURIComponent(firstLine)
      return `https://www.linkedin.com/jobs/search/?keywords=${q}`
    }
    if (cUp.includes('SALARY')) {
      return `https://www.levels.fyi/?search=${encodeURIComponent(firstLine)}`
    }
    return `https://www.linkedin.com/in/${slug}`
  }

  return ''
}

export interface MergedCellLayout {
  spanMap: Map<string, number>
  hiddenSet: Set<string>
  mergedGroupCount: number
  mergedRowCount: number
  firstMergedOriginalIndex: number | null
}

/**
 * Detects vertically merged rows in a sheet (both in Existing Data and Newly Produced Data).
 * A merged block starts at row `i` when the primary column (column 0, e.g. NAME) is populated
 * and 1 or more subsequent rows `i+1 .. j-1` have an empty primary column `""` while having
 * values in secondary columns (such as COMPANY, JOB TYPE, POSITION), or share `_group_id`.
 */
function computeMergedCellLayout(
  indexedRows: { row: SheetRow; originalIndex: number }[],
  headers: string[]
): MergedCellLayout {
  const spanMap = new Map<string, number>()
  const hiddenSet = new Set<string>()
  let mergedGroupCount = 0
  let mergedRowCount = 0
  let firstMergedOriginalIndex: number | null = null

  if (!indexedRows.length || !headers.length) {
    return { spanMap, hiddenSet, mergedGroupCount, mergedRowCount, firstMergedOriginalIndex }
  }

  const primaryHeader = headers[0]
  const secondaryHeaders = headers.slice(1)

  let i = 0
  while (i < indexedRows.length) {
    const startRow = indexedRows[i].row
    const primaryVal = String(getCellValue(startRow, primaryHeader) ?? '').trim()
    const startGroupId =
      typeof startRow._group_id === 'string' ? startRow._group_id : ''

    // Skip instructional banner rows
    const isBannerRow = primaryVal.length > 90 && primaryVal.includes('Ex.')

    if (!primaryVal || isBannerRow) {
      i += 1
      continue
    }

    let j = i + 1
    while (j < indexedRows.length) {
      const nextEntry = indexedRows[j]
      const nextRow = nextEntry.row
      // Must be consecutive in the underlying sheet (or share explicit _group_id)
      const isConsecutive =
        nextEntry.originalIndex === indexedRows[j - 1].originalIndex + 1
      const nextGroupId =
        typeof nextRow._group_id === 'string' ? nextRow._group_id : ''

      if (startGroupId && nextGroupId && startGroupId === nextGroupId) {
        j += 1
        continue
      }

      if (!isConsecutive) break

      const nextPrimaryVal = String(getCellValue(nextRow, primaryHeader) ?? '').trim()
      const hasSecondaryVal = secondaryHeaders.some(
        (h) => String(getCellValue(nextRow, h) ?? '').trim() !== ''
      )

      if (nextPrimaryVal === '' && hasSecondaryVal) {
        j += 1
      } else {
        break
      }
    }

    const span = j - i
    if (span > 1) {
      mergedGroupCount += 1
      mergedRowCount += span
      if (firstMergedOriginalIndex === null) {
        firstMergedOriginalIndex = indexedRows[i].originalIndex
      }

      // Any column that has a non-empty value on the top row `i` and is blank `""`
      // on all continuation rows `i+1 .. j-1` is vertically merged across the group!
      for (const h of headers) {
        const topVal = String(getCellValue(startRow, h) ?? '').trim()
        if (!topVal) continue
        let allContinuationBlank = true
        for (let k = i + 1; k < j; k += 1) {
          if (String(getCellValue(indexedRows[k].row, h) ?? '').trim() !== '') {
            allContinuationBlank = false
            break
          }
        }
        if (allContinuationBlank) {
          spanMap.set(`${i}:${h}`, span)
          for (let k = i + 1; k < j; k += 1) {
            hiddenSet.add(`${k}:${h}`)
          }
        }
      }
      i = j
    } else {
      i += 1
    }
  }

  return { spanMap, hiddenSet, mergedGroupCount, mergedRowCount, firstMergedOriginalIndex }
}

/**
 * Filters rows by search query while keeping all rows of a merged multi-row entity block together.
 */
function filterRowsPreservingMergedGroups(
  rawRows: SheetRow[],
  headers: string[],
  searchQuery: string
): { row: SheetRow; originalIndex: number }[] {
  const allIndexed = rawRows.map((row, idx) => ({ row, originalIndex: idx }))
  const q = searchQuery.trim().toLowerCase()
  if (!q) return allIndexed

  const primaryHeader = headers[0] || ''
  const secondaryHeaders = headers.slice(1)

  // Group consecutive merged rows first
  const groups: { row: SheetRow; originalIndex: number }[][] = []
  for (let i = 0; i < allIndexed.length; i += 1) {
    const item = allIndexed[i]
    const pVal = primaryHeader
      ? String(getCellValue(item.row, primaryHeader) ?? '').trim()
      : ''
    const gid = typeof item.row._group_id === 'string' ? item.row._group_id : ''
    const hasSecondary = secondaryHeaders.some(
      (h) => String(getCellValue(item.row, h) ?? '').trim() !== ''
    )

    if (
      groups.length > 0 &&
      ((gid && gid === groups[groups.length - 1][0].row._group_id) ||
        (primaryHeader && pVal === '' && hasSecondary))
    ) {
      groups[groups.length - 1].push(item)
    } else {
      groups.push([item])
    }
  }

  const matched: { row: SheetRow; originalIndex: number }[] = []
  for (const grp of groups) {
    const groupMatches = grp.some(({ row }) =>
      Object.entries(row).some(
        ([k, v]) => !k.startsWith('_') && String(v ?? '').toLowerCase().includes(q)
      )
    )
    if (groupMatches) {
      matched.push(...grp)
    }
  }
  return matched
}

function buildGoogleSheetsTsv(
  rows: SheetRow[],
  headers: string[],
  columns: ColumnTemplate[] = [],
  includeHeaders = false,
  useHyperlinkFormulas = false
): string {
  const colMap = new Map<string, ColumnTemplate>()
  for (const c of columns) {
    colMap.set(c.name.trim().toUpperCase(), c)
  }

  const lines: string[] = []
  if (includeHeaders && headers.length > 0) {
    lines.push(headers.map((h) => String(h).replace(/[\t\r\n]+/g, ' ').trim()).join('\t'))
  }
  for (const row of rows) {
    const cells = headers.map((h) => {
      const colTemplate = colMap.get(h.trim().toUpperCase())
      const rawVal = String(getCellValue(row, h) ?? '')
        .replace(/\t/g, ' ')
        .replace(/\r\n/g, '\n')
        .replace(/\r/g, '\n')
        .trim()
      const linkUrl = getCellLink(row, h, colTemplate)

      if (useHyperlinkFormulas && linkUrl && rawVal && !rawVal.includes('\n')) {
        const safeUrl = linkUrl.replace(/"/g, '""')
        const safeLabel = rawVal.replace(/"/g, '""')
        return `=HYPERLINK("${safeUrl}", "${safeLabel}")`
      }

      // Quote multi-line cells (e.g. bullet points `• Apple\n• Microsoft\n• Google`)
      // so Google Sheets pastes all bullet points into a single cell!
      if (rawVal.includes('\n') || rawVal.includes('"')) {
        return `"${rawVal.replace(/"/g, '""')}"`
      }
      return rawVal
    })
    lines.push(cells.join('\t'))
  }
  return lines.join('\n')
}

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function buildGoogleSheetsHtmlTable(
  rows: SheetRow[],
  headers: string[],
  columns: ColumnTemplate[] = [],
  includeHeaders = false
): string {
  const colMap = new Map<string, ColumnTemplate>()
  for (const c of columns) {
    colMap.set(c.name.trim().toUpperCase(), c)
  }

  const indexed = rows.map((row, idx) => ({ row, originalIndex: idx }))
  const { spanMap, hiddenSet } = computeMergedCellLayout(indexed, headers)

  const trs: string[] = []
  if (includeHeaders && headers.length > 0) {
    const ths = headers.map((h) => `<th>${escapeHtml(String(h))}</th>`).join('')
    trs.push(`<tr>${ths}</tr>`)
  }

  rows.forEach((row, rIdx) => {
    const tds: string[] = []
    for (const h of headers) {
      const cellKey = `${rIdx}:${h}`
      if (hiddenSet.has(cellKey)) {
        // Covered by a vertically merged `<td rowspan="N">` on an earlier row
        continue
      }
      const span = spanMap.get(cellKey) || 1
      const colTemplate = colMap.get(h.trim().toUpperCase())
      const rawVal = String(getCellValue(row, h) ?? '').trim()
      const linkUrl = getCellLink(row, h, colTemplate)

      // Convert internal newlines (bullet points) to <br> for single-cell multi-line paste in Google Sheets
      const escapedLines = rawVal
        .split('\n')
        .map((line) => escapeHtml(line))
        .join('<br>')

      const innerHtml =
        linkUrl && rawVal
          ? `<a href="${escapeHtml(linkUrl)}">${escapedLines}</a>`
          : escapedLines

      const rowSpanAttr = span > 1 ? ` rowspan="${span}"` : ''
      const styleAttr =
        rawVal.includes('\n') || span > 1
          ? ` style="${span > 1 ? 'vertical-align:middle;' : ''}${
              rawVal.includes('\n') ? 'white-space:pre-wrap;' : ''
            }"`
          : ''

      tds.push(`<td${rowSpanAttr}${styleAttr}>${innerHtml}</td>`)
    }
    trs.push(`<tr>${tds.join('')}</tr>`)
  })

  return `<table><tbody>${trs.join('')}</tbody></table>`
}

interface SheetCellProps {
  rowIndex: number
  colName: string
  value: string | number
  linkUrl: string
  span: number
  onCommit: (rowIndex: number, colName: string, newValue: string) => void
}

/**
 * Memoized cell component with isolated local draft state.
 * Prevents keystrokes inside a cell (or in any other input on the page) from
 * re-rendering all 800–1,600 table cells on every character typed.
 */
const MemoizedSheetCell = React.memo(function MemoizedSheetCell({
  rowIndex,
  colName,
  value,
  linkUrl,
  span,
  onCommit,
}: SheetCellProps) {
  const externalStr = String(value ?? '')
  const [draft, setDraft] = useState<string>(externalStr)

  useEffect(() => {
    setDraft(String(value ?? ''))
  }, [value])

  const isMultiLine = draft.includes('\n') || externalStr.includes('\n')

  function commitIfChanged(currentVal: string) {
    if (currentVal !== externalStr) {
      onCommit(rowIndex, colName, currentVal)
    }
  }

  return (
    <td
      rowSpan={span > 1 ? span : undefined}
      className={span > 1 ? 'dm-merged-cell' : undefined}
    >
      <div className="dm-cell-inner">
        {isMultiLine ? (
          <textarea
            className="dm-sheet-cell-textarea"
            rows={Math.min(5, Math.max(2, draft.split('\n').length))}
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            onBlur={(e) => commitIfChanged(e.target.value)}
          />
        ) : (
          <input
            type="text"
            className={`dm-sheet-cell-input ${linkUrl ? 'dm-linked-cell-input' : ''}`}
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            onBlur={(e) => commitIfChanged(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.currentTarget.blur()
              }
            }}
          />
        )}
        {(linkUrl || span > 1) && (
          <div className="dm-cell-badges">
            {linkUrl && (
              <a
                href={linkUrl}
                target="_blank"
                rel="noopener noreferrer"
                className="dm-cell-link-badge"
                title={`Open webpage: ${linkUrl}`}
                onClick={(e) => e.stopPropagation()}
              >
                🔗 Link ↗
              </a>
            )}
            {span > 1 && (
              <span
                className="dm-cell-merged-badge"
                title={`Vertically merged across ${span} rows`}
              >
                ↕ {span} rows
              </span>
            )}
          </div>
        )}
      </div>
    </td>
  )
})

export function App() {
  // Auth state
  const [user, setUser] = useState<User | null>(null)
  const [authLoading, setAuthLoading] = useState(true)
  const [authMode, setAuthMode] = useState<'login' | 'signup'>('login')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [authError, setAuthError] = useState('')
  const [authMessage, setAuthMessage] = useState('')
  const [submittingAuth, setSubmittingAuth] = useState(false)

  // Projects & Navigation state
  const [projects, setProjects] = useState<Project[]>([])
  const [loadingProjects, setLoadingProjects] = useState(false)
  const [activeProject, setActiveProject] = useState<Project | null>(null)
  const [selectedColumnId, setSelectedColumnId] = useState<string>('')

  // New project creation state
  const [creatingNewProject, setCreatingNewProject] = useState(false)
  const [newProjectName, setNewProjectName] = useState('')

  // Workspace UI state
  const [savingStatus, setSavingStatus] = useState<
    'idle' | 'saving' | 'saved' | 'syncing-gsheet' | 'synced-gsheet' | 'error'
  >('idle')
  const [statusError, setStatusError] = useState('')
  const [fetchingData, setFetchingData] = useState(false)
  const [fetchProgress, setFetchProgress] = useState<FetchProgress | null>(null)
  const [agentPrompt, setAgentPrompt] = useState('')
  const [agentMode, setAgentMode] = useState<
    'gemini-3.8-flash' | 'antigravity-agent' | 'gemini-4-argon'
  >('gemini-3.8-flash')
  const [requestedLinesInput, setRequestedLinesInput] = useState<string>('25')
  const [addingColumnName, setAddingColumnName] = useState('')

  // New Data (Google Sheets Ready) state
  const [newDataPage, setNewDataPage] = useState(1)
  const [newDataSearch, setNewDataSearch] = useState('')
  const deferredNewDataSearch = useDeferredValue(newDataSearch)
  const [copyToast, setCopyToast] = useState('')
  const [showRawTsvPreview, setShowRawTsvPreview] = useState(false)

  // Tab & Table Pagination/Search state
  const [tablePage, setTablePage] = useState(1)
  const [tableSearch, setTableSearch] = useState('')
  const deferredTableSearch = useDeferredValue(tableSearch)
  const [renamingTabId, setRenamingTabId] = useState<string | null>(null)
  const [renamingTabValue, setRenamingTabValue] = useState('')

  // Modals: Add MCP & AI Settings
  const [showAddMcpModal, setShowAddMcpModal] = useState(false)
  const [newMcpName, setNewMcpName] = useState('')
  const [newMcpUrl, setNewMcpUrl] = useState('')
  const [showAiKeyModal, setShowAiKeyModal] = useState(false)
  const [geminiApiKey, setGeminiApiKey] = useState<string>(
    () => localStorage.getItem('DRS_GEMINI_API_KEY') || ''
  )
  const [customApiBaseUrl, setCustomApiBaseUrl] = useState<string>(
    () => localStorage.getItem('DRS_API_BASE_URL') || ''
  )

  const fileInputRef = useRef<HTMLInputElement | null>(null)
  const activeProjectRef = useRef<Project | null>(null)
  useEffect(() => {
    activeProjectRef.current = activeProject
  }, [activeProject])

  // Observe Firebase Auth state
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (firebaseUser) => {
      setUser(firebaseUser)
      setAuthLoading(false)
    })
    return () => unsub()
  }, [])

  // Load projects when authenticated
  useEffect(() => {
    if (!user) {
      setProjects([])
      setActiveProject(null)
      return
    }
    loadUserProjects()
  }, [user])

  async function loadUserProjects() {
    setLoadingProjects(true)
    setStatusError('')
    try {
      const list = await fetchProjects()
      setProjects(list)
    } catch (err: unknown) {
      setStatusError(err instanceof Error ? err.message : 'Failed to load projects')
    } finally {
      setLoadingProjects(false)
    }
  }

  // Handle Login / Sign Up
  async function handleAuthSubmit(e: React.FormEvent) {
    e.preventDefault()
    setAuthError('')
    setAuthMessage('')
    setSubmittingAuth(true)
    try {
      if (authMode === 'login') {
        await signInWithEmailAndPassword(auth, email.trim(), password)
      } else {
        await createUserWithEmailAndPassword(auth, email.trim(), password)
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Authentication failed'
      setAuthError(msg.replace('Firebase: ', ''))
    } finally {
      setSubmittingAuth(false)
    }
  }

  async function handlePasswordReset() {
    if (!email.trim()) {
      setAuthError('Enter your email address above first to receive a password reset link.')
      return
    }
    setAuthError('')
    try {
      await sendPasswordResetEmail(auth, email.trim())
      setAuthMessage(`Password reset email sent to ${email.trim()}.`)
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Failed to send password reset email'
      setAuthError(msg.replace('Firebase: ', ''))
    }
  }

  // Open a project into Workspace view
  function openProject(proj: Project) {
    setActiveProject(proj)
    if (proj.columns && proj.columns.length > 0) {
      setSelectedColumnId(proj.columns[0].id)
    } else {
      setSelectedColumnId('')
    }
    setTablePage(1)
    setTableSearch('')
    setStatusError('')
  }

  // Persist project updates WITHOUT overwriting live local input state when the network response arrives
  async function persistProjectUpdates(
    nextProject: Project,
    fields: Partial<
      Pick<
        Project,
        | 'name'
        | 'existing_data_source'
        | 'existing_data_type'
        | 'active_tab_id'
        | 'tabs'
        | 'columns'
        | 'mcps'
        | 'new_rows'
      >
    > & { sync_to_gsheet?: boolean }
  ) {
    setActiveProject(nextProject)
    setProjects((prev) => prev.map((p) => (p.id === nextProject.id ? nextProject : p)))
    setSavingStatus(fields.sync_to_gsheet ? 'syncing-gsheet' : 'saving')
    try {
      const saved = await updateProject(nextProject.id, fields)
      // Update only server metadata (`last_edited` and `gsheet_sync`) so in-flight typing is never clobbered
      setActiveProject((curr) =>
        curr && curr.id === saved.id
          ? {
              ...curr,
              last_edited: saved.last_edited || curr.last_edited,
              gsheet_sync: saved.gsheet_sync || curr.gsheet_sync,
            }
          : curr
      )
      setProjects((prev) =>
        prev.map((p) =>
          p.id === saved.id
            ? {
                ...p,
                last_edited: saved.last_edited || p.last_edited,
                gsheet_sync: saved.gsheet_sync || p.gsheet_sync,
              }
            : p
        )
      )
      setSavingStatus(saved.gsheet_sync?.synced ? 'synced-gsheet' : 'saved')
    } catch (err: unknown) {
      setSavingStatus('error')
      setStatusError(err instanceof Error ? err.message : 'Failed to save changes')
    }
  }

  // Create a new project
  async function handleCreateProject(e: React.FormEvent) {
    e.preventDefault()
    const trimmed = newProjectName.trim() || `Project-${projects.length + 1}`
    try {
      const created = await createProject(trimmed)
      setProjects((prev) => [created, ...prev])
      setNewProjectName('')
      setCreatingNewProject(false)
      openProject(created)
    } catch (err: unknown) {
      setStatusError(err instanceof Error ? err.message : 'Could not create project')
    }
  }

  // Delete a project
  async function handleDeleteProject(e: React.MouseEvent, projectId: string) {
    e.stopPropagation()
    try {
      await deleteProject(projectId)
      setProjects((prev) => prev.filter((p) => p.id !== projectId))
      if (activeProject?.id === projectId) {
        setActiveProject(null)
      }
    } catch (err: unknown) {
      setStatusError(err instanceof Error ? err.message : 'Could not delete project')
    }
  }

  // CSV Upload handler
  async function handleCsvFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file || !activeProject) return
    setSavingStatus('saving')
    setStatusError('')
    try {
      const updated = await uploadCsvFile(
        activeProject.id,
        file,
        activeProject.active_tab_id,
        true
      )
      setActiveProject(updated)
      setProjects((prev) => prev.map((p) => (p.id === updated.id ? updated : p)))
      if (
        updated.columns.length > 0 &&
        !updated.columns.some((c) => c.id === selectedColumnId)
      ) {
        setSelectedColumnId(updated.columns[0].id)
      }
      setTablePage(1)
      setSavingStatus('saved')
    } catch (err: unknown) {
      setSavingStatus('error')
      setStatusError(err instanceof Error ? err.message : 'Failed to upload CSV')
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = ''
    }
  }

  // Google Sheets Import & Multi-Tab Sync handler
  async function handleGoogleSheetsImport(explicitUrl?: string) {
    if (!activeProject) return
    const currentVal = (explicitUrl ?? activeProject.existing_data_source).trim()
    const sheetUrl = currentVal.includes('docs.google.com/spreadsheets')
      ? currentVal
      : window.prompt(
          'Paste your Google Sheets URL to sync all tabs, columns, and rows:',
          currentVal || 'https://docs.google.com/spreadsheets/d/'
        ) || ''

    if (!sheetUrl.trim()) return
    setSavingStatus('syncing-gsheet')
    setStatusError('')
    try {
      const updated = await importGoogleSheet(
        activeProject.id,
        sheetUrl.trim(),
        activeProject.active_tab_id,
        true
      )
      setActiveProject(updated)
      setProjects((prev) => prev.map((p) => (p.id === updated.id ? updated : p)))
      if (updated.columns.length > 0) {
        setSelectedColumnId(updated.columns[0].id)
      }
      setTablePage(1)
      setTableSearch('')
      setSavingStatus('synced-gsheet')
    } catch (err: unknown) {
      setSavingStatus('error')
      setStatusError(err instanceof Error ? err.message : 'Failed to sync Google Sheet')
    }
  }

  // Explicit Pull / Push Google Sheet Sync
  async function handleSyncGoogleSheet(direction: 'pull' | 'push') {
    if (!activeProject) return
    setSavingStatus('syncing-gsheet')
    setStatusError('')
    try {
      const updated = await syncGoogleSheet(
        activeProject.id,
        direction,
        activeProject.active_tab_id
      )
      setActiveProject(updated)
      setProjects((prev) => prev.map((p) => (p.id === updated.id ? updated : p)))
      if (
        updated.columns.length > 0 &&
        !updated.columns.some((c) => c.id === selectedColumnId)
      ) {
        setSelectedColumnId(updated.columns[0].id)
      }
      setSavingStatus('synced-gsheet')
    } catch (err: unknown) {
      setSavingStatus('error')
      setStatusError(err instanceof Error ? err.message : `Failed to ${direction} Google Sheet`)
    }
  }

  // Update a specific ColumnTemplate property in local state ONLY (for fast, zero-lag typing in text inputs)
  function updateSelectedColumnLocal(patch: Partial<ColumnTemplate>) {
    if (!activeProject) return
    const nextCols = activeProject.columns.map((c) => {
      if (c.id === selectedColumnId) {
        return { ...c, ...patch }
      }
      if (patch.retrieve_all_mode) {
        return { ...c, retrieve_all_mode: patch.retrieve_all_mode }
      }
      return c
    })
    setActiveProject({ ...activeProject, columns: nextCols })
  }

  // Persist current columns to backend on blur of Column Name / Special Instructions / Filter Value
  function commitColumnsToBackend() {
    if (!activeProject) return
    persistProjectUpdates(activeProject, { columns: activeProject.columns })
  }

  // Update a specific ColumnTemplate property AND persist immediately (for discrete button/checkbox/select clicks)
  function updateSelectedColumnImmediate(patch: Partial<ColumnTemplate>) {
    if (!activeProject) return
    const nextCols = activeProject.columns.map((c) => {
      if (c.id === selectedColumnId) {
        return { ...c, ...patch }
      }
      // Keep retrieve_all_mode consistent across all columns when toggled
      if (patch.retrieve_all_mode) {
        return { ...c, retrieve_all_mode: patch.retrieve_all_mode }
      }
      return c
    })
    const nextProj = { ...activeProject, columns: nextCols }
    persistProjectUpdates(nextProj, { columns: nextCols })
  }

  // Add a new column to Data Template (and active tab headers)
  function handleAddColumn(e: React.FormEvent) {
    e.preventDefault()
    if (!activeProject) return
    const name = addingColumnName.trim()
    if (!name) return
    const newCol: ColumnTemplate = {
      id: `col-${Date.now()}`,
      name,
      special_instructions: '',
      sort: null,
      query_type: 'Attribute',
      data_type: 'String',
      filter_op: 'None',
      filter_value: '',
    }
    const nextCols = [...activeProject.columns, newCol]
    const nextTabs = activeProject.tabs.map((t) => {
      if (t.name.trim().toUpperCase().startsWith('RULE')) return t
      const hdrs = t.headers && t.headers.length > 0 ? [...t.headers] : activeProject.columns.map((c) => c.name)
      if (!hdrs.includes(name)) hdrs.push(name)
      return { ...t, headers: hdrs }
    })
    const nextProj = { ...activeProject, columns: nextCols, tabs: nextTabs }
    setAddingColumnName('')
    setSelectedColumnId(newCol.id)
    persistProjectUpdates(nextProj, { columns: nextCols, tabs: nextTabs })
  }

  // Delete currently selected column
  function handleDeleteColumn(colId: string) {
    if (!activeProject || activeProject.columns.length <= 1) return
    const targetCol = activeProject.columns.find((c) => c.id === colId)
    const nextCols = activeProject.columns.filter((c) => c.id !== colId)
    const nextTabs = activeProject.tabs.map((t) => {
      if (!targetCol || !t.headers) return t
      return { ...t, headers: t.headers.filter((h) => h !== targetCol.name) }
    })
    const nextProj = { ...activeProject, columns: nextCols, tabs: nextTabs }
    setSelectedColumnId(nextCols[0]?.id || '')
    persistProjectUpdates(nextProj, { columns: nextCols, tabs: nextTabs })
  }

  // Sheet Tab actions (Instant local switch + background persistence)
  function handleSwitchTab(tabId: string) {
    if (!activeProject || activeProject.active_tab_id === tabId) return
    const nextProj = { ...activeProject, active_tab_id: tabId }
    setActiveProject(nextProj)
    setProjects((prev) => prev.map((p) => (p.id === nextProj.id ? nextProj : p)))
    setTablePage(1)
    setTableSearch('')
    // Save active_tab_id in background without sending huge tabs payload
    updateProject(activeProject.id, { active_tab_id: tabId }).catch(() => {})
  }

  function handleAddTab() {
    if (!activeProject) return
    const nextIndex = activeProject.tabs.length + 1
    const defaultHeaders = activeProject.columns.map((c) => c.name)
    const newTab = {
      id: `tab-${Date.now()}`,
      name: `Tab ${nextIndex}`,
      headers: defaultHeaders,
      rows: [],
    }
    const nextTabs = [...activeProject.tabs, newTab]
    const nextProj = { ...activeProject, tabs: nextTabs, active_tab_id: newTab.id }
    setTablePage(1)
    setTableSearch('')
    persistProjectUpdates(nextProj, {
      tabs: nextTabs,
      active_tab_id: newTab.id,
      sync_to_gsheet: activeProject.existing_data_type === 'gsheets',
    })
  }

  function handleCommitTabRename(tabId: string) {
    if (!activeProject) return
    const trimmed = renamingTabValue.trim()
    setRenamingTabId(null)
    if (!trimmed) return
    const nextTabs = activeProject.tabs.map((t) =>
      t.id === tabId ? { ...t, name: trimmed } : t
    )
    const nextProj = { ...activeProject, tabs: nextTabs }
    persistProjectUpdates(nextProj, {
      tabs: nextTabs,
      sync_to_gsheet: activeProject.existing_data_type === 'gsheets',
    })
  }

  function handleDeleteTab(e: React.MouseEvent, tabId: string) {
    e.stopPropagation()
    if (!activeProject || activeProject.tabs.length <= 1) return
    const nextTabs = activeProject.tabs.filter((t) => t.id !== tabId)
    const nextActiveId =
      activeProject.active_tab_id === tabId ? nextTabs[0].id : activeProject.active_tab_id
    const nextProj = { ...activeProject, tabs: nextTabs, active_tab_id: nextActiveId }
    setTablePage(1)
    persistProjectUpdates(nextProj, { tabs: nextTabs, active_tab_id: nextActiveId })
  }

  // Commit a cell edit in the Existing Data Sheet table (called on blur / Enter from MemoizedSheetCell)
  const handleCommitExistingCell = useCallback(
    async (actualRowIndex: number, colName: string, value: string) => {
      const proj = activeProjectRef.current
      if (!proj) return
      const nextTabs = proj.tabs.map((tab) => {
        if (tab.id !== proj.active_tab_id) return tab
        const nextRows = [...tab.rows]
        nextRows[actualRowIndex] = { ...nextRows[actualRowIndex], [colName]: value }
        return { ...tab, rows: nextRows }
      })
      const nextProj = { ...proj, tabs: nextTabs }
      setActiveProject(nextProj)

      const isGsheet = proj.existing_data_type === 'gsheets'
      setSavingStatus(isGsheet ? 'syncing-gsheet' : 'saving')
      try {
        const res = await updateSheetCell(
          proj.id,
          proj.active_tab_id,
          actualRowIndex,
          colName,
          value,
          isGsheet
        )
        setProjects((prev) => prev.map((p) => (p.id === nextProj.id ? nextProj : p)))
        setSavingStatus(res.gsheet_sync?.synced ? 'synced-gsheet' : 'saved')
      } catch (err: unknown) {
        setSavingStatus('error')
        setStatusError(err instanceof Error ? err.message : 'Failed to save cell')
      }
    },
    []
  )

  async function handleAddEmptySheetRow() {
    if (!activeProject) return
    const isGsheet = activeProject.existing_data_type === 'gsheets'
    setSavingStatus(isGsheet ? 'syncing-gsheet' : 'saving')
    try {
      const updated = await addRowToTab(
        activeProject.id,
        activeProject.active_tab_id,
        undefined,
        isGsheet
      )
      setActiveProject(updated)
      setProjects((prev) => prev.map((p) => (p.id === updated.id ? updated : p)))
      const updatedTab = updated.tabs.find((t) => t.id === updated.active_tab_id)
      if (updatedTab) {
        setTablePage(Math.max(1, Math.ceil(updatedTab.rows.length / PAGE_SIZE)))
      }
      setSavingStatus(updated.gsheet_sync?.synced ? 'synced-gsheet' : 'saved')
    } catch (err: unknown) {
      setSavingStatus('error')
      setStatusError(err instanceof Error ? err.message : 'Failed to add row')
    }
  }

  function handleExportCsv() {
    if (!activeProject) return
    const activeTab =
      activeProject.tabs.find((t) => t.id === activeProject.active_tab_id) ||
      activeProject.tabs[0]
    if (!activeTab) return
    const headers =
      activeTab.headers && activeTab.headers.length > 0
        ? activeTab.headers
        : activeProject.columns.map((c) => c.name)
    const csvLines = [
      headers.map((h) => `"${h.replace(/"/g, '""')}"`).join(','),
      ...activeTab.rows.map((row) =>
        headers
          .map((h) => {
            const val = getCellValue(row, h)
            return `"${String(val).replace(/"/g, '""')}"`
          })
          .join(',')
      ),
    ]
    const blob = new Blob([csvLines.join('\n')], { type: 'text/csv;charset=utf-8;' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${activeProject.name}-${activeTab.name}.csv`
    a.click()
    URL.revokeObjectURL(url)
  }

  // Toggle MCP connector
  function handleToggleMcp(mcpId: string) {
    if (!activeProject) return
    const nextMcps = activeProject.mcps.map((m) =>
      m.id === mcpId ? { ...m, enabled: !m.enabled } : m
    )
    const nextProj = { ...activeProject, mcps: nextMcps }
    persistProjectUpdates(nextProj, { mcps: nextMcps })
  }

  // Add custom MCP connector
  function handleCreateCustomMcp(e: React.FormEvent) {
    e.preventDefault()
    if (!activeProject || !newMcpName.trim()) return
    const customMcp: MCPConnector = {
      id: `mcp-${Date.now()}`,
      name: newMcpName.trim(),
      icon: 'custom',
      enabled: true,
      server_url: newMcpUrl.trim(),
      description: 'Custom MCP / Web Scraping Source',
    }
    const nextMcps = [...activeProject.mcps, customMcp]
    const nextProj = { ...activeProject, mcps: nextMcps }
    setNewMcpName('')
    setNewMcpUrl('')
    setShowAddMcpModal(false)
    persistProjectUpdates(nextProj, { mcps: nextMcps })
  }

  // Run Agentic AI "Fetch Data" (Produces ONLY new non-repeating data, up to 1000 lines)
  async function handleFetchData() {
    if (!activeProject) return
    const parsedLines = parseInt(requestedLinesInput, 10)
    const clampedLines = Number.isNaN(parsedLines)
      ? 10
      : Math.max(1, Math.min(1000, parsedLines))
    setRequestedLinesInput(String(clampedLines))
    const projectId = activeProject.id
    const runId = crypto.randomUUID().slice(0, 12)
    // A new request replaces the previous new data; rows from this run stream in below.
    setActiveProject((prev) =>
      prev && prev.id === projectId ? { ...prev, new_rows: [], last_agent_run: null } : prev
    )
    setFetchProgress(null)
    setNewDataPage(1)
    setNewDataSearch('')
    setFetchingData(true)
    setStatusError('')
    setCopyToast('')

    let polling = true
    let rowsSeen = 0
    const pollProgress = async () => {
      while (polling) {
        try {
          const res = await getFetchProgress(projectId, rowsSeen)
          // Ignore state left over from a previous run until the backend has started this one.
          if (polling && res.progress?.run_id === runId) {
            const freshRows = res.rows
            rowsSeen = res.total_rows
            setFetchProgress(res.progress)
            if (freshRows.length > 0) {
              setActiveProject((prev) =>
                prev && prev.id === projectId
                  ? { ...prev, new_rows: [...(prev.new_rows || []), ...freshRows] }
                  : prev
              )
            }
          }
        } catch {
          // Transient polling errors are fine; the final response carries the full result.
        }
        await new Promise((resolve) => setTimeout(resolve, 2000))
      }
    }
    void pollProgress()

    try {
      await updateProject(activeProject.id, {
        name: activeProject.name,
        existing_data_source: activeProject.existing_data_source,
        existing_data_type: activeProject.existing_data_type,
        active_tab_id: activeProject.active_tab_id,
        columns: activeProject.columns,
        mcps: activeProject.mcps,
      })
      const updated = await runFetchDataAgent(activeProject.id, {
        query_prompt: agentPrompt,
        tab_id: activeProject.active_tab_id,
        row_count: clampedLines,
        append_rows: false,
        agent_mode: agentMode,
        gemini_api_key: geminiApiKey || undefined,
        run_id: runId,
      })
      polling = false
      setActiveProject(updated)
      setProjects((prev) => prev.map((p) => (p.id === updated.id ? updated : p)))
      setNewDataPage(1)
      setSavingStatus('saved')
    } catch (err: unknown) {
      setStatusError(err instanceof Error ? err.message : 'Failed to fetch data with AI Agent')
    } finally {
      polling = false
      setFetchingData(false)
      setFetchProgress(null)
    }
  }

  async function handleClearNewData() {
    if (!activeProject) return
    if (!window.confirm('Clear all new data produced by the last fetch? This cannot be undone.')) return
    setStatusError('')
    try {
      const updated = await clearNewData(activeProject.id)
      setActiveProject(updated)
      setProjects((prev) => prev.map((p) => (p.id === updated.id ? updated : p)))
      setNewDataPage(1)
      setNewDataSearch('')
    } catch (err: unknown) {
      setStatusError(err instanceof Error ? err.message : 'Failed to clear new data')
    }
  }

  // Copy Newly Produced Data to Clipboard for Google Sheets (TSV + HTML Table with Links, Bullet Points & Merged Rows)
  async function handleCopyNewDataToClipboard(
    rowsToCopy: SheetRow[],
    headersToUse: string[],
    includeHeaders: boolean,
    useHyperlinkFormulas = false
  ) {
    if (!rowsToCopy.length) return
    const cols = activeProject?.columns || []
    const tsvText = buildGoogleSheetsTsv(
      rowsToCopy,
      headersToUse,
      cols,
      includeHeaders,
      useHyperlinkFormulas
    )
    const htmlTable = buildGoogleSheetsHtmlTable(rowsToCopy, headersToUse, cols, includeHeaders)

    let copiedOk = false
    try {
      if (
        !useHyperlinkFormulas &&
        typeof ClipboardItem !== 'undefined' &&
        navigator.clipboard?.write
      ) {
        const item = new ClipboardItem({
          'text/plain': new Blob([tsvText], { type: 'text/plain' }),
          'text/html': new Blob([htmlTable], { type: 'text/html' }),
        })
        await navigator.clipboard.write([item])
        copiedOk = true
      }
    } catch {
      copiedOk = false
    }

    if (!copiedOk) {
      try {
        await navigator.clipboard.writeText(tsvText)
        copiedOk = true
      } catch {
        // Fallback textarea copy
        const ta = document.createElement('textarea')
        ta.value = tsvText
        ta.style.position = 'fixed'
        ta.style.opacity = '0'
        document.body.appendChild(ta)
        ta.focus()
        ta.select()
        try {
          document.execCommand('copy')
          copiedOk = true
        } finally {
          document.body.removeChild(ta)
        }
      }
    }

    if (copiedOk) {
      setCopyToast(
        `✓ Copied ${rowsToCopy.length.toLocaleString()} new ${
          rowsToCopy.length === 1 ? 'row' : 'rows'
        } (${
          useHyperlinkFormulas
            ? 'with =HYPERLINK() formulas'
            : includeHeaders
            ? 'with column headers, links, bullets & merged cells'
            : 'rows only — with clickable links, bullet points & merged cells'
        }) to clipboard! Open Google Sheets and press Cmd+V (or Ctrl+V) to paste.`
      )
      setTimeout(() => setCopyToast(''), 6000)
    }
  }

  // Commit a cell edit inside the Newly Produced Data table (called on blur / Enter from MemoizedSheetCell)
  const handleCommitNewDataCell = useCallback(
    (rowIndex: number, colName: string, value: string) => {
      const proj = activeProjectRef.current
      if (!proj || !proj.new_rows) return
      const nextNewRows = [...proj.new_rows]
      nextNewRows[rowIndex] = { ...nextNewRows[rowIndex], [colName]: value }
      const nextProj = { ...proj, new_rows: nextNewRows }
      persistProjectUpdates(nextProj, { new_rows: nextNewRows })
    },
    []
  )

  function handleExportNewDataCsv(
    rowsToExport: SheetRow[],
    headersToUse: string[]
  ) {
    if (!activeProject || !rowsToExport.length) return
    const csvLines = [
      headersToUse.map((h) => `"${h.replace(/"/g, '""')}"`).join(','),
      ...rowsToExport.map((row) =>
        headersToUse
          .map((h) => {
            const val = getCellValue(row, h)
            return `"${String(val).replace(/"/g, '""')}"`
          })
          .join(',')
      ),
    ]
    const blob = new Blob([csvLines.join('\n')], { type: 'text/csv;charset=utf-8;' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${activeProject.name}-new-data-${rowsToExport.length}-rows.csv`
    a.click()
    URL.revokeObjectURL(url)
  }

  // Save Gemini API Key & Optional Custom Backend API URL in localStorage
  function handleSaveApiKey(e: React.FormEvent) {
    e.preventDefault()
    localStorage.setItem('DRS_GEMINI_API_KEY', geminiApiKey.trim())
    const trimmedUrl = customApiBaseUrl.trim().replace(/\/+$/, '')
    if (trimmedUrl) {
      localStorage.setItem('DRS_API_BASE_URL', trimmedUrl)
    } else {
      localStorage.removeItem('DRS_API_BASE_URL')
    }
    setCustomApiBaseUrl(trimmedUrl)
    setShowAiKeyModal(false)
  }

  // Compute active tab, headers, filtered rows, and current page rows with narrowed dependencies
  const projectTabs = activeProject?.tabs
  const activeTabId = activeProject?.active_tab_id
  const projectColumns = activeProject?.columns

  const activeTab = useMemo(() => {
    if (!projectTabs?.length) return null
    return projectTabs.find((t) => t.id === activeTabId) || projectTabs[0]
  }, [projectTabs, activeTabId])

  const displayHeaders = useMemo(() => {
    if (!projectColumns) return []
    if (activeTab?.headers && activeTab.headers.length > 0) {
      return activeTab.headers
    }
    const templateNames = projectColumns.map((c) => c.name).filter(Boolean)
    if (activeTab?.rows && activeTab.rows.length > 0) {
      const sampleKeys = Object.keys(activeTab.rows[0]).filter((k) => !k.startsWith('_'))
      // Check if template names match sampleKeys (case-insensitively)
      const hasMatch = templateNames.some((tn) =>
        sampleKeys.some((sk) => sk.trim().toUpperCase() === tn.trim().toUpperCase())
      )
      if (!hasMatch && sampleKeys.length > 0) {
        return sampleKeys
      }
    }
    return templateNames
  }, [projectColumns, activeTab])

  const activeTabRows = activeTab?.rows
  const indexedFilteredRows = useMemo(() => {
    const rawRows = activeTabRows || []
    return filterRowsPreservingMergedGroups(rawRows, displayHeaders, deferredTableSearch)
  }, [activeTabRows, displayHeaders, deferredTableSearch])

  const existingTabMergedStats = useMemo(() => {
    const rawRows = activeTabRows || []
    const allIndexed = rawRows.map((row, idx) => ({ row, originalIndex: idx }))
    return computeMergedCellLayout(allIndexed, displayHeaders)
  }, [activeTabRows, displayHeaders])

  const totalPages = Math.max(1, Math.ceil(indexedFilteredRows.length / PAGE_SIZE))
  const safePage = Math.min(tablePage, totalPages)
  const visibleRows = useMemo(() => {
    const start = (safePage - 1) * PAGE_SIZE
    return indexedFilteredRows.slice(start, start + PAGE_SIZE)
  }, [indexedFilteredRows, safePage])

  const visibleMergedLayout = useMemo(
    () => computeMergedCellLayout(visibleRows, displayHeaders),
    [visibleRows, displayHeaders]
  )

  // Newly Produced Data (Only New Data, never repeating Existing Data, formatted for Google Sheets)
  const rawNewRows = activeProject?.new_rows
  const rawNewHeaders = activeProject?.new_headers
  const newRows = useMemo(() => rawNewRows || [], [rawNewRows])
  const newDataHeaders = useMemo(() => {
    if (!projectColumns) return []
    if (rawNewHeaders && rawNewHeaders.length > 0) {
      return rawNewHeaders
    }
    if (
      activeTab &&
      !activeTab.name.trim().toUpperCase().startsWith('RULE') &&
      activeTab.headers &&
      activeTab.headers.length > 0
    ) {
      return activeTab.headers
    }
    return projectColumns.map((c) => c.name).filter(Boolean)
  }, [projectColumns, rawNewHeaders, activeTab])

  const indexedFilteredNewRows = useMemo(() => {
    return filterRowsPreservingMergedGroups(newRows, newDataHeaders, deferredNewDataSearch)
  }, [newRows, newDataHeaders, deferredNewDataSearch])

  const newRowsMergedStats = useMemo(() => {
    const allIndexed = newRows.map((row, idx) => ({ row, originalIndex: idx }))
    return computeMergedCellLayout(allIndexed, newDataHeaders)
  }, [newRows, newDataHeaders])

  const totalNewPages = Math.max(1, Math.ceil(indexedFilteredNewRows.length / PAGE_SIZE))
  const safeNewPage = Math.min(newDataPage, totalNewPages)
  const visibleNewRows = useMemo(() => {
    const start = (safeNewPage - 1) * PAGE_SIZE
    return indexedFilteredNewRows.slice(start, start + PAGE_SIZE)
  }, [indexedFilteredNewRows, safeNewPage])

  const visibleNewMergedLayout = useMemo(
    () => computeMergedCellLayout(visibleNewRows, newDataHeaders),
    [visibleNewRows, newDataHeaders]
  )

  // ---------------------------------------------------------------------------
  // Render Loading State
  // ---------------------------------------------------------------------------
  if (authLoading) {
    return (
      <div className="dm-page">
        <header className="dm-header">
          <h1 className="dm-title">Data Mining</h1>
        </header>
        <main className="dm-main">
          <p style={{ color: '#666' }}>Initializing Firebase Authentication...</p>
        </main>
      </div>
    )
  }

  // ---------------------------------------------------------------------------
  // Render Password-Protected Login / Sign Up Gate
  // ---------------------------------------------------------------------------
  if (!user) {
    return (
      <div className="dm-page">
        <header className="dm-header">
          <div className="dm-header-left">
            <h1 className="dm-title">Data Mining</h1>
          </div>
        </header>
        <main className="dm-main">
          <div className="dm-auth-wrap">
            <h2 className="dm-auth-title">
              {authMode === 'login' ? 'Sign In to Data Mining' : 'Create Account'}
            </h2>
            <p className="dm-auth-subtitle">
              Password-protected workspace connected to Firebase Authentication &amp; Cloud
              Firestore.
            </p>

            {authError && <div className="dm-error-banner">{authError}</div>}
            {authMessage && (
              <div
                className="dm-error-banner"
                style={{ background: '#ecfdf5', borderColor: '#a7f3d0', color: '#065f46' }}
              >
                {authMessage}
              </div>
            )}

            <form onSubmit={handleAuthSubmit}>
              <div className="dm-auth-field">
                <label htmlFor="dm-email">Email Address</label>
                <input
                  id="dm-email"
                  type="email"
                  required
                  placeholder="you@company.com"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                />
              </div>

              <div className="dm-auth-field">
                <label htmlFor="dm-password">Password</label>
                <input
                  id="dm-password"
                  type="password"
                  required
                  minLength={6}
                  placeholder="••••••••"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                />
              </div>

              <button type="submit" className="dm-auth-submit" disabled={submittingAuth}>
                {submittingAuth
                  ? 'Authenticating...'
                  : authMode === 'login'
                  ? 'Sign In'
                  : 'Create Account'}
              </button>
            </form>

            <div className="dm-auth-switch">
              {authMode === 'login' ? (
                <>
                  Need an account?
                  <button
                    type="button"
                    className="dm-auth-link"
                    onClick={() => {
                      setAuthMode('signup')
                      setAuthError('')
                    }}
                  >
                    Sign Up
                  </button>
                  {' · '}
                  <button type="button" className="dm-auth-link" onClick={handlePasswordReset}>
                    Reset Password
                  </button>
                </>
              ) : (
                <>
                  Already have an account?
                  <button
                    type="button"
                    className="dm-auth-link"
                    onClick={() => {
                      setAuthMode('login')
                      setAuthError('')
                    }}
                  >
                    Sign In
                  </button>
                </>
              )}
            </div>
          </div>
        </main>
      </div>
    )
  }

  // ---------------------------------------------------------------------------
  // Render Screen 1: Projects Dashboard (Matches Screenshot 1)
  // ---------------------------------------------------------------------------
  if (!activeProject) {
    const totalRows = Math.max(6, projects.length)
    const emptyRowCount = totalRows - projects.length

    return (
      <div className="dm-page">
        <header className="dm-header">
          <div className="dm-header-left">
            <h1 className="dm-title">Data Mining</h1>
          </div>
          <div className="dm-header-right">
            <span className="dm-user-pill">{user.email}</span>
            <button
              type="button"
              className="dm-header-btn"
              onClick={() => setShowAiKeyModal(true)}
            >
              AI Config ({agentMode === 'gemini-3.8-flash' ? 'Gemini 3.8 Flash' : agentMode})
            </button>
            <button type="button" className="dm-header-btn" onClick={() => signOut(auth)}>
              Sign Out
            </button>
          </div>
        </header>

        <main className="dm-main">
          {statusError && <div className="dm-error-banner">{statusError}</div>}

          <div className="dm-new-project-bar">
            <button
              type="button"
              className="dm-new-project-btn"
              onClick={() => setCreatingNewProject((prev) => !prev)}
            >
              New Project +
            </button>

            {creatingNewProject && (
              <form className="dm-new-project-inline" onSubmit={handleCreateProject}>
                <input
                  type="text"
                  placeholder="Enter project name (e.g. Lead-outreach)"
                  value={newProjectName}
                  onChange={(e) => setNewProjectName(e.target.value)}
                  autoFocus
                />
                <button
                  type="submit"
                  className="dm-header-btn"
                  style={{ background: '#ffa000', borderColor: '#ffa000' }}
                >
                  Create
                </button>
                <button
                  type="button"
                  className="dm-header-btn"
                  onClick={() => setCreatingNewProject(false)}
                >
                  Cancel
                </button>
              </form>
            )}
          </div>

          <div className="dm-projects-table-wrap">
            <div className="dm-projects-header">
              <div>Name</div>
              <div style={{ textAlign: 'center' }}>Last Edited</div>
              <div></div>
            </div>

            {loadingProjects ? (
              <div style={{ padding: '24px 32px', color: '#666' }}>
                Loading projects from Cloud Firestore...
              </div>
            ) : (
              <div className="dm-projects-list">
                {projects.map((proj, idx) => (
                  <div
                    key={proj.id}
                    className={`dm-project-row clickable ${idx % 2 === 0 ? 'even' : 'odd'}`}
                    onClick={() => openProject(proj)}
                  >
                    <div>{proj.name}</div>
                    <div className="dm-project-date" style={{ textAlign: 'center' }}>
                      {formatLastEdited(proj.last_edited)}
                    </div>
                    <button
                      type="button"
                      className="dm-project-delete-btn"
                      title="Delete project"
                      onClick={(e) => handleDeleteProject(e, proj.id)}
                    >
                      ✕
                    </button>
                  </div>
                ))}

                {Array.from({ length: emptyRowCount }).map((_, i) => {
                  const rowIdx = projects.length + i
                  return (
                    <div
                      key={`empty-${rowIdx}`}
                      className={`dm-project-row empty-row ${rowIdx % 2 === 0 ? 'even' : 'odd'}`}
                    >
                      <div></div>
                      <div></div>
                      <div></div>
                    </div>
                  )
                })}
              </div>
            )}
          </div>
        </main>

        {showAiKeyModal && (
          <div className="dm-modal-backdrop" onClick={() => setShowAiKeyModal(false)}>
            <div className="dm-modal" onClick={(e) => e.stopPropagation()}>
              <h3>Agentic AI Configuration</h3>
              <p className="dm-auth-subtitle">
                We selected <strong>Gemini 3.8 Flash</strong> (with native Google Search grounding,
                URL Context scraping, MCP support, and JSON Schema structured output) +{' '}
                <strong>Antigravity Managed Agent</strong> as the best AI stack for sheet-style web
                scraping.
              </p>
              <form onSubmit={handleSaveApiKey}>
                <div className="dm-auth-field">
                  <label>Default Agentic Model</label>
                  <select
                    className="dm-agent-select"
                    style={{ width: '100%' }}
                    value={agentMode}
                    onChange={(e) =>
                      setAgentMode(
                        e.target.value as
                          | 'gemini-3.8-flash'
                          | 'antigravity-agent'
                          | 'gemini-4-argon'
                      )
                    }
                  >
                    <option value="gemini-3.8-flash">
                      Gemini 3.8 Flash (Fast Search + Structured Sheet Output)
                    </option>
                    <option value="antigravity-agent">
                      Antigravity Agent (antigravity-preview-09-2026 Sandbox)
                    </option>
                    <option value="gemini-4-argon">
                      Gemini 4 Argon (Deep Reasoning &amp; Entity Extraction)
                    </option>
                  </select>
                </div>
                <div className="dm-auth-field">
                  <label>Gemini API Key (optional if set in backend/.env)</label>
                  <input
                    type="password"
                    placeholder="Configured in backend/.env"
                    value={geminiApiKey}
                    onChange={(e) => setGeminiApiKey(e.target.value)}
                  />
                </div>
                <div className="dm-auth-field">
                  <label>Backend API Base URL (optional override for Vercel / remote backend)</label>
                  <input
                    type="text"
                    placeholder={
                      import.meta.env.VITE_API_BASE_URL || 'http://127.0.0.1:8000'
                    }
                    value={customApiBaseUrl}
                    onChange={(e) => setCustomApiBaseUrl(e.target.value)}
                  />
                </div>
                <div style={{ display: 'flex', gap: '10px', justifyContent: 'flex-end' }}>
                  <button
                    type="button"
                    className="dm-header-btn"
                    onClick={() => setShowAiKeyModal(false)}
                  >
                    Close
                  </button>
                  <button
                    type="submit"
                    className="dm-header-btn"
                    style={{ background: '#ffa000', borderColor: '#ffa000' }}
                  >
                    Save Settings
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}
      </div>
    )
  }

  // ---------------------------------------------------------------------------
  // Render Screen 2: Project Workspace & Agentic Data Mining View (Matches Screenshot 2)
  // ---------------------------------------------------------------------------
  const selectedCol =
    activeProject.columns.find((c) => c.id === selectedColumnId) || activeProject.columns[0]
  const isGoogleSheetLinked =
    activeProject.existing_data_type === 'gsheets' ||
    activeProject.existing_data_source.includes('docs.google.com/spreadsheets')

  return (
    <div className="dm-page">
      <header className="dm-header">
        <div className="dm-header-left">
          <h1
            className="dm-title"
            title="Return to Projects List"
            onClick={() => setActiveProject(null)}
          >
            Data Mining
          </h1>
          <div className="dm-project-subtitle">
            <span>/</span>
            <input
              type="text"
              className="dm-project-name-input"
              value={activeProject.name}
              onChange={(e) => setActiveProject({ ...activeProject, name: e.target.value })}
              onBlur={() =>
                persistProjectUpdates(activeProject, { name: activeProject.name })
              }
            />
          </div>
        </div>

        <div className="dm-header-right">
          <button
            type="button"
            className="dm-header-btn"
            onClick={() => setActiveProject(null)}
          >
            ← All Projects
          </button>
          <button
            type="button"
            className="dm-header-btn"
            onClick={() => setShowAiKeyModal(true)}
          >
            AI: {agentMode === 'gemini-3.8-flash' ? 'Gemini 3.8 Flash' : agentMode}
          </button>
          <span
            className="dm-user-pill"
            style={
              savingStatus === 'synced-gsheet'
                ? { background: '#ecfdf5', borderColor: '#a7f3d0', color: '#065f46' }
                : undefined
            }
          >
            {savingStatus === 'syncing-gsheet'
              ? 'Syncing with Google Sheets...'
              : savingStatus === 'synced-gsheet'
              ? 'Synced with Google Sheets ✓'
              : savingStatus === 'saving'
              ? 'Saving to Firestore...'
              : savingStatus === 'saved'
              ? 'Saved ✓'
              : user.email}
          </span>
          <button type="button" className="dm-header-btn" onClick={() => signOut(auth)}>
            Sign Out
          </button>
        </div>
      </header>

      <main className="dm-main">
        {statusError && <div className="dm-error-banner">{statusError}</div>}

        {/* 1. EXISTING DATA SECTION */}
        <section className="dm-existing-data-section">
          <div className="dm-section-label">
            <span>Existing Data</span>
            {isGoogleSheetLinked && (
              <span
                style={{
                  fontFamily: 'sans-serif',
                  fontSize: '0.78rem',
                  color: '#107c41',
                  fontWeight: 600,
                }}
              >
                ● Live Google Sheet Connected ({activeProject.tabs.length} tabs)
              </span>
            )}
          </div>

          <div className="dm-existing-box">
            <div className="dm-existing-input-row">
              <input
                type="text"
                className="dm-existing-input"
                placeholder="Paste a Google Sheets URL and press Enter (or click Google Sheets below), or upload a .csv file..."
                value={activeProject.existing_data_source}
                onChange={(e) =>
                  setActiveProject({ ...activeProject, existing_data_source: e.target.value })
                }
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault()
                    if (
                      activeProject.existing_data_source.includes(
                        'docs.google.com/spreadsheets'
                      )
                    ) {
                      handleGoogleSheetsImport(activeProject.existing_data_source)
                    } else {
                      persistProjectUpdates(activeProject, {
                        existing_data_source: activeProject.existing_data_source,
                      })
                    }
                  }
                }}
                onBlur={() =>
                  persistProjectUpdates(activeProject, {
                    existing_data_source: activeProject.existing_data_source,
                  })
                }
              />
              <button
                type="button"
                className="dm-existing-clear"
                title="Clear existing data source"
                onClick={() => {
                  const nextProj = {
                    ...activeProject,
                    existing_data_source: '',
                    existing_data_type: 'none' as const,
                  }
                  persistProjectUpdates(nextProj, {
                    existing_data_source: '',
                    existing_data_type: 'none',
                  })
                }}
              >
                ×
              </button>
            </div>

            {/* Multi-Tab Bar (Supports all Google Sheet worksheets, renaming, and adding tabs) */}
            <div className="dm-tabs-bar">
              {activeProject.tabs.map((tab) => {
                const isActive = tab.id === (activeTab?.id || activeProject.active_tab_id)
                return (
                  <button
                    key={tab.id}
                    type="button"
                    className={`dm-tab-btn ${isActive ? 'active' : ''}`}
                    onClick={() => handleSwitchTab(tab.id)}
                    onDoubleClick={() => {
                      setRenamingTabId(tab.id)
                      setRenamingTabValue(tab.name)
                    }}
                    title="Click to switch tab · Double-click to rename"
                  >
                    {renamingTabId === tab.id ? (
                      <input
                        type="text"
                        value={renamingTabValue}
                        onChange={(e) => setRenamingTabValue(e.target.value)}
                        onBlur={() => handleCommitTabRename(tab.id)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') handleCommitTabRename(tab.id)
                          if (e.key === 'Escape') setRenamingTabId(null)
                        }}
                        onClick={(e) => e.stopPropagation()}
                        autoFocus
                        style={{ width: '90px', fontSize: '0.75rem', padding: '1px 4px' }}
                      />
                    ) : (
                      <>
                        <span>{tab.name}</span>
                        <span className="dm-tab-count">{tab.rows?.length || 0}</span>
                        {activeProject.tabs.length > 1 && (
                          <span
                            className="dm-tab-close"
                            title="Delete tab"
                            onClick={(e) => handleDeleteTab(e, tab.id)}
                          >
                            ×
                          </span>
                        )}
                      </>
                    )}
                  </button>
                )
              })}
              <button
                type="button"
                className="dm-tab-btn add-tab"
                title="Add new worksheet tab"
                onClick={handleAddTab}
              >
                +
              </button>
            </div>
          </div>

          {/* File (.csv) and Google Sheets Buttons */}
          <div className="dm-source-buttons">
            <input
              ref={fileInputRef}
              type="file"
              accept=".csv,text/csv"
              style={{ display: 'none' }}
              onChange={handleCsvFileChange}
            />
            <button
              type="button"
              className={`dm-source-btn ${
                activeProject.existing_data_type === 'csv' ? 'active-source' : ''
              }`}
              onClick={() => fileInputRef.current?.click()}
            >
              <span style={{ fontSize: '0.9rem' }}>📄</span>
              <span>File (.csv)</span>
            </button>

            <button
              type="button"
              className={`dm-source-btn ${isGoogleSheetLinked ? 'active-source' : ''}`}
              onClick={() => handleGoogleSheetsImport()}
            >
              <span className="gsheet-badge">⊞</span>
              <span>Google Sheets</span>
            </button>

            {isGoogleSheetLinked && (
              <>
                <button
                  type="button"
                  className="dm-source-btn"
                  title="Pull latest tabs, columns, and rows from the Google Sheet"
                  onClick={() => handleSyncGoogleSheet('pull')}
                >
                  <span>↻ Pull from Sheet</span>
                </button>
                <button
                  type="button"
                  className="dm-source-btn"
                  title="Push current tab rows & edits directly to the Google Sheet"
                  onClick={() => handleSyncGoogleSheet('push')}
                >
                  <span>↑ Push to Sheet</span>
                </button>
              </>
            )}
          </div>

          {/* Interactive Sheet-Style Format Data Grid */}
          <div className="dm-sheet-container">
            <div className="dm-sheet-toolbar">
              <div className="dm-sheet-toolbar-left">
                <strong>{activeTab?.name || 'Tab 1'}</strong>
                <span style={{ color: '#666' }}>
                  ({indexedFilteredRows.length.toLocaleString()}{' '}
                  {indexedFilteredRows.length === 1 ? 'row' : 'rows'})
                </span>
                {existingTabMergedStats.mergedGroupCount > 0 && (
                  <span
                    className="dm-merged-detected-pill"
                    title={`Detected ${existingTabMergedStats.mergedGroupCount} vertically merged multi-row groups (${existingTabMergedStats.mergedRowCount} rows) in ${activeTab?.name}`}
                  >
                    ⊞ {existingTabMergedStats.mergedGroupCount} Merged Row{' '}
                    {existingTabMergedStats.mergedGroupCount === 1 ? 'Group' : 'Groups'} Detected
                  </span>
                )}
                {existingTabMergedStats.mergedGroupCount > 0 && (
                  <button
                    type="button"
                    className="dm-header-btn dm-jump-merged-btn"
                    title="Jump to merged rows (e.g. Colby Philbin on the last page or first merged group)"
                    onClick={() => {
                      setTableSearch('')
                      // Jump to last page where Colby Philbin's 5-row merged block is located, or toggle to first merged block
                      if (safePage === totalPages && existingTabMergedStats.firstMergedOriginalIndex !== null) {
                        const firstMergedPage =
                          Math.floor(existingTabMergedStats.firstMergedOriginalIndex / PAGE_SIZE) + 1
                        setTablePage(firstMergedPage)
                      } else {
                        setTablePage(totalPages)
                      }
                    }}
                  >
                    {safePage === totalPages && existingTabMergedStats.firstMergedOriginalIndex !== null
                      ? '↱ Jump to First Merged Group'
                      : '↳ Jump to Last Merged Group (Colby Philbin)'}
                  </button>
                )}
                <input
                  type="text"
                  className="dm-sheet-search"
                  placeholder={`Search in ${activeTab?.name || 'sheet'}...`}
                  value={tableSearch}
                  onChange={(e) => {
                    setTableSearch(e.target.value)
                    setTablePage(1)
                  }}
                />
              </div>

              <div className="dm-sheet-toolbar-right">
                <button
                  type="button"
                  className="dm-header-btn"
                  onClick={handleAddEmptySheetRow}
                >
                  + Add Row
                </button>
                <button type="button" className="dm-header-btn" onClick={handleExportCsv}>
                  Export .CSV
                </button>
              </div>
            </div>

            <div className="dm-sheet-scroll">
              <table className="dm-sheet-table">
                <thead>
                  <tr>
                    <th className="dm-sheet-row-num">#</th>
                    {displayHeaders.map((headerName) => {
                      const matchingCol = activeProject.columns.find(
                        (c) => c.name.trim().toUpperCase() === headerName.trim().toUpperCase()
                      )
                      return (
                        <th
                          key={headerName}
                          style={{ cursor: matchingCol ? 'pointer' : 'default' }}
                          onClick={() => {
                            if (matchingCol) setSelectedColumnId(matchingCol.id)
                          }}
                          title={
                            matchingCol?.special_instructions ||
                            'Click to configure column in Data Template'
                          }
                        >
                          {headerName}
                          {matchingCol && (
                            <span
                              style={{
                                fontWeight: 400,
                                fontSize: '0.7rem',
                                color: '#666',
                                marginLeft: '5px',
                              }}
                            >
                              ({matchingCol.data_type}
                              {matchingCol.sort
                                ? ` • ${matchingCol.sort === 'Ascending' ? '↑' : '↓'}`
                                : ''}
                              {matchingCol.put_link ? ' • 🔗' : ''}
                              {matchingCol.retrieve_all ? ' • All' : ''}
                              )
                            </span>
                          )}
                        </th>
                      )
                    })}
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.length === 0 ? (
                    <tr>
                      <td
                        colSpan={displayHeaders.length + 1}
                        style={{
                          padding: '22px',
                          textAlign: 'center',
                          color: '#777',
                        }}
                      >
                        {tableSearch
                          ? `No rows match "${tableSearch}" in ${activeTab?.name}.`
                          : `No rows in ${activeTab?.name} yet. Click Fetch Data below to scrape data using Agentic AI, or sync a Google Sheet / CSV.`}
                      </td>
                    </tr>
                  ) : (
                    visibleRows.map(({ row, originalIndex }, visIdx) => (
                      <tr key={originalIndex}>
                        <td className="dm-sheet-row-num">{originalIndex + 1}</td>
                        {displayHeaders.map((headerName) => {
                          const cellKey = `${visIdx}:${headerName}`
                          if (visibleMergedLayout.hiddenSet.has(cellKey)) {
                            // Covered by a vertically merged cell on an earlier row
                            return null
                          }
                          const span = visibleMergedLayout.spanMap.get(cellKey) || 1
                          const matchingCol = activeProject.columns.find(
                            (c) =>
                              c.name.trim().toUpperCase() === headerName.trim().toUpperCase()
                          )
                          const cellVal = getCellValue(row, headerName)
                          const linkUrl = getCellLink(row, headerName, matchingCol)

                          return (
                            <MemoizedSheetCell
                              key={headerName}
                              rowIndex={originalIndex}
                              colName={headerName}
                              value={cellVal}
                              linkUrl={linkUrl}
                              span={span}
                              onCommit={handleCommitExistingCell}
                            />
                          )
                        })}
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination Bar for large sheets (e.g. 1,030+ rows) */}
            {indexedFilteredRows.length > PAGE_SIZE && (
              <div className="dm-sheet-pagination">
                <div>
                  Showing {(safePage - 1) * PAGE_SIZE + 1}–
                  {Math.min(safePage * PAGE_SIZE, indexedFilteredRows.length)} of{' '}
                  {indexedFilteredRows.length.toLocaleString()} rows in{' '}
                  <strong>{activeTab?.name}</strong>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <button
                    type="button"
                    className="dm-header-btn"
                    disabled={safePage <= 1}
                    onClick={() => setTablePage(1)}
                  >
                    « First
                  </button>
                  <button
                    type="button"
                    className="dm-header-btn"
                    disabled={safePage <= 1}
                    onClick={() => setTablePage((p) => Math.max(1, p - 1))}
                  >
                    ‹ Prev
                  </button>
                  <span style={{ padding: '0 8px' }}>
                    Page <strong>{safePage}</strong> of <strong>{totalPages}</strong>
                  </span>
                  <button
                    type="button"
                    className="dm-header-btn"
                    disabled={safePage >= totalPages}
                    onClick={() => setTablePage((p) => Math.min(totalPages, p + 1))}
                  >
                    Next ›
                  </button>
                  <button
                    type="button"
                    className="dm-header-btn"
                    disabled={safePage >= totalPages}
                    onClick={() => setTablePage(totalPages)}
                  >
                    Last »
                  </button>
                </div>
              </div>
            )}
          </div>
        </section>

        {/* 2. DATA TEMPLATE SECTION (Matches Screenshot 2) */}
        <section className="dm-template-section">
          <div className="dm-section-label">
            <span>Data Template</span>
          </div>

          <div className="dm-template-box">
            <div className="dm-template-header">
              <div>Column</div>
              <div>Properties</div>
            </div>

            <div className="dm-template-rows">
              {activeProject.columns.map((col, idx) => {
                const isSelected = col.id === selectedCol?.id
                const propSummary = [
                  col.query_type,
                  col.data_type,
                  col.sort ? `Sort: ${col.sort}` : null,
                  col.put_link ? '🔗 Put Link' : null,
                  col.retrieve_all
                    ? `Retrieve All (${
                        col.retrieve_all_mode === 'more_rows'
                          ? 'More Rows / Merged'
                          : 'Bullet Points'
                      })`
                    : null,
                  col.filter_op !== 'None' && col.filter_value
                    ? `Filter: ${col.filter_op} ${col.filter_value}`
                    : null,
                  col.special_instructions
                    ? `"${
                        col.special_instructions.length > 65
                          ? col.special_instructions.slice(0, 65) + '...'
                          : col.special_instructions
                      }"`
                    : null,
                ]
                  .filter(Boolean)
                  .join(' · ')

                return (
                  <div
                    key={col.id}
                    className={`dm-template-row ${idx % 2 === 0 ? 'even' : 'odd'} ${
                      isSelected ? 'selected-col' : ''
                    }`}
                    onClick={() => setSelectedColumnId(col.id)}
                  >
                    <div className="dm-template-col-cell">{col.name}</div>
                    <div className="dm-template-prop-cell">
                      <span>{propSummary}</span>
                      <span className="dm-chevron">{isSelected ? '⌄' : '›'}</span>
                    </div>
                  </div>
                )
              })}

              {/* Extra alternating row to add new columns cleanly */}
              <div
                className={`dm-template-row ${
                  activeProject.columns.length % 2 === 0 ? 'even' : 'odd'
                }`}
              >
                <div className="dm-template-col-cell">
                  <form
                    onSubmit={handleAddColumn}
                    style={{ display: 'flex', width: '100%' }}
                  >
                    <input
                      type="text"
                      placeholder="+ Add column..."
                      value={addingColumnName}
                      onChange={(e) => setAddingColumnName(e.target.value)}
                      style={{
                        border: 'none',
                        background: 'transparent',
                        width: '100%',
                        fontSize: '0.86rem',
                        outline: 'none',
                      }}
                    />
                  </form>
                </div>
                <div className="dm-template-prop-cell">
                  {addingColumnName.trim() && (
                    <button
                      type="button"
                      className="dm-header-btn"
                      style={{ padding: '2px 8px', fontSize: '0.75rem' }}
                      onClick={handleAddColumn}
                    >
                      Add Column
                    </button>
                  )}
                </div>
              </div>

              <div
                className={`dm-template-row ${
                  (activeProject.columns.length + 1) % 2 === 0 ? 'even' : 'odd'
                }`}
              >
                <div className="dm-template-col-cell"></div>
                <div className="dm-template-prop-cell"></div>
              </div>
            </div>

            {/* Expanded Column Editor Panel (Column: Name) */}
            {selectedCol && (
              <div className="dm-column-editor">
                <div className="dm-column-editor-header">
                  <div className="dm-column-heading">
                    <span>Column:</span>
                    <input
                      type="text"
                      className="dm-column-rename-input"
                      value={selectedCol.name}
                      onChange={(e) => updateSelectedColumnLocal({ name: e.target.value })}
                      onBlur={commitColumnsToBackend}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          e.currentTarget.blur()
                        }
                      }}
                    />
                  </div>
                  {activeProject.columns.length > 1 && (
                    <button
                      type="button"
                      className="dm-header-btn"
                      onClick={() => handleDeleteColumn(selectedCol.id)}
                    >
                      Remove Column
                    </button>
                  )}
                </div>

                <div className="dm-special-instructions-wrap">
                  <label className="dm-special-label">Special Instructions</label>
                  <textarea
                    className="dm-special-input"
                    placeholder="Instructions for the AI agent when scraping this column (automatically populated from RULES! tab if present)..."
                    value={selectedCol.special_instructions}
                    onChange={(e) =>
                      updateSelectedColumnLocal({ special_instructions: e.target.value })
                    }
                    onBlur={commitColumnsToBackend}
                  />
                </div>

                <div className="dm-controls-grid">
                  {/* Line 1: Sort & Query Type */}
                  <div className="dm-control-line">
                    <div className="dm-control-group">
                      <span className="dm-control-label">Sort:</span>
                      <div className="dm-segmented">
                        {(['Ascending', 'Descending'] as SortDirection[]).map((dir) => (
                          <button
                            key={dir}
                            type="button"
                            className={`dm-seg-btn ${
                              selectedCol.sort === dir ? 'active' : ''
                            }`}
                            onClick={() =>
                              updateSelectedColumnImmediate({
                                sort: selectedCol.sort === dir ? null : dir,
                              })
                            }
                          >
                            {dir}
                          </button>
                        ))}
                      </div>
                    </div>

                    <div className="dm-control-group">
                      <span className="dm-control-label">Query Type:</span>
                      <div className="dm-segmented">
                        {(['Main', 'Attribute'] as QueryType[]).map((qt) => (
                          <button
                            key={qt}
                            type="button"
                            className={`dm-seg-btn ${
                              selectedCol.query_type === qt ? 'active' : ''
                            }`}
                            onClick={() => updateSelectedColumnImmediate({ query_type: qt })}
                          >
                            {qt}
                          </button>
                        ))}
                      </div>
                    </div>
                  </div>

                  {/* Line 2: Data Type */}
                  <div className="dm-control-line">
                    <div className="dm-control-group">
                      <span className="dm-control-label">Data Type:</span>
                      <div className="dm-segmented">
                        {(['String', 'Number', 'Date'] as DataType[]).map((dt) => (
                          <button
                            key={dt}
                            type="button"
                            className={`dm-seg-btn ${
                              selectedCol.data_type === dt ? 'active' : ''
                            }`}
                            onClick={() => updateSelectedColumnImmediate({ data_type: dt })}
                          >
                            {dt}
                          </button>
                        ))}
                      </div>
                    </div>
                  </div>

                  {/* Line 3: Filter */}
                  <div className="dm-control-line">
                    <div className="dm-control-group">
                      <span className="dm-control-label">Filter:</span>
                      <select
                        className="dm-filter-select"
                        value={selectedCol.filter_op}
                        onChange={(e) =>
                          updateSelectedColumnImmediate({
                            filter_op: e.target.value as FilterOp,
                          })
                        }
                      >
                        <option value="Less Than">Less Than</option>
                        <option value="Greater Than">Greater Than</option>
                        <option value="Equals">Equals</option>
                        <option value="Contains">Contains</option>
                        <option value="None">None</option>
                      </select>

                      <input
                        type="text"
                        className="dm-filter-input"
                        placeholder="Value..."
                        value={selectedCol.filter_value}
                        onChange={(e) =>
                          updateSelectedColumnLocal({ filter_value: e.target.value })
                        }
                        onBlur={commitColumnsToBackend}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') {
                            e.currentTarget.blur()
                          }
                        }}
                      />
                    </div>
                  </div>

                  {/* Line 4: Put Link & Retrieve All Information Checkboxes */}
                  <div className="dm-control-line dm-checkbox-line">
                    <label
                      className="dm-checkbox-card"
                      title="Inserts the person/entity's profile or company webpage URL as a clickable link that copies and pastes directly into Google Sheets"
                    >
                      <input
                        type="checkbox"
                        checked={Boolean(selectedCol.put_link)}
                        onChange={(e) =>
                          updateSelectedColumnImmediate({ put_link: e.target.checked })
                        }
                      />
                      <div className="dm-checkbox-text">
                        <span className="dm-checkbox-title">Put link</span>
                        <span className="dm-checkbox-desc">
                          Attaches webpage filled with information (pastes as a clickable link in
                          Google Sheets)
                        </span>
                      </div>
                    </label>

                    <label
                      className="dm-checkbox-card"
                      title="Retrieves all historical/multi-entry data for this sector (e.g. worked at Apple, Microsoft, Google, etc.)"
                    >
                      <input
                        type="checkbox"
                        checked={Boolean(selectedCol.retrieve_all)}
                        onChange={(e) =>
                          updateSelectedColumnImmediate({
                            retrieve_all: e.target.checked,
                            retrieve_all_mode:
                              selectedCol.retrieve_all_mode ||
                              (existingTabMergedStats.mergedGroupCount > 0
                                ? 'more_rows'
                                : 'bullet_points'),
                          })
                        }
                      />
                      <div className="dm-checkbox-text">
                        <span className="dm-checkbox-title">Retrieve all information</span>
                        <span className="dm-checkbox-desc">
                          Retrieves all data for this sector (e.g., worked for Apple, Microsoft,
                          Google, etc.)
                        </span>
                      </div>
                    </label>
                  </div>

                  {/* Line 5: Format Selector when Retrieve all information is enabled */}
                  {selectedCol.retrieve_all && (
                    <div className="dm-control-line dm-retrieve-mode-box">
                      <div className="dm-control-group" style={{ flexWrap: 'wrap' }}>
                        <span className="dm-control-label" style={{ minWidth: '140px', fontWeight: 600 }}>
                          Retrieve All Format:
                        </span>
                        <div className="dm-segmented">
                          {(
                            [
                              {
                                mode: 'bullet_points' as RetrieveAllMode,
                                label: '• Insert through Bullet Points',
                              },
                              {
                                mode: 'more_rows' as RetrieveAllMode,
                                label: '⊞ Insert through More Rows (Merged)',
                              },
                            ] as const
                          ).map((item) => {
                            const currentMode =
                              selectedCol.retrieve_all_mode || 'bullet_points'
                            return (
                              <button
                                key={item.mode}
                                type="button"
                                className={`dm-seg-btn ${
                                  currentMode === item.mode ? 'active' : ''
                                }`}
                                onClick={() =>
                                  updateSelectedColumnImmediate({
                                    retrieve_all_mode: item.mode,
                                  })
                                }
                              >
                                {item.label}
                              </button>
                            )
                          })}
                        </div>
                        <span className="dm-retrieve-mode-hint">
                          {(selectedCol.retrieve_all_mode || 'bullet_points') ===
                          'bullet_points'
                            ? 'Single row per person — formats all entries as • Apple / • Microsoft / • Google inside one cell (quoted for Google Sheets paste).'
                            : 'Multiple rows per person — merges identity cells vertically across rows while listing each company/role on its own row (matches Google Sheets merged rows).'}
                        </span>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        </section>

        {/* 3. RECOMMENDED MCP'S SECTION (Matches Screenshot 2) */}
        <section className="dm-mcps-section">
          <div className="dm-section-label">
            <span>Recommended MCP’s</span>
          </div>

          <div className="dm-mcps-row">
            {activeProject.mcps.map((mcp) => {
              if (mcp.icon === 'linkedin') {
                return (
                  <div
                    key={mcp.id}
                    className={`dm-mcp-card dm-mcp-linkedin ${
                      mcp.enabled ? 'active-mcp' : 'disabled-mcp'
                    }`}
                    title={`${mcp.name} MCP Connector (${
                      mcp.enabled ? 'Active' : 'Click to enable'
                    })`}
                    onClick={() => handleToggleMcp(mcp.id)}
                  >
                    <span>in</span>
                    {mcp.enabled && <span className="dm-mcp-check">✓</span>}
                  </div>
                )
              }
              if (mcp.icon === 'handshake') {
                return (
                  <div
                    key={mcp.id}
                    className={`dm-mcp-card dm-mcp-handshake ${
                      mcp.enabled ? 'active-mcp' : 'disabled-mcp'
                    }`}
                    title={`${mcp.name} MCP Connector (${
                      mcp.enabled ? 'Active' : 'Click to enable'
                    })`}
                    onClick={() => handleToggleMcp(mcp.id)}
                  >
                    <span>H</span>
                    {mcp.enabled && <span className="dm-mcp-check">✓</span>}
                  </div>
                )
              }
              return (
                <div
                  key={mcp.id}
                  className={`dm-mcp-card dm-mcp-custom ${
                    mcp.enabled ? 'active-mcp' : 'disabled-mcp'
                  }`}
                  title={`${mcp.name} (${mcp.server_url || 'Custom Connector'})`}
                  onClick={() => handleToggleMcp(mcp.id)}
                >
                  <span>{mcp.name}</span>
                  {mcp.enabled && <span className="dm-mcp-check">✓</span>}
                </div>
              )
            })}

            <button
              type="button"
              className="dm-mcp-card dm-mcp-add"
              title="Add Custom MCP Server or Data Source"
              onClick={() => setShowAddMcpModal(true)}
            >
              +
            </button>
          </div>
        </section>

        {/* 4. AGENTIC AI PROMPT, LINE COUNT (1-1000) & FETCH DATA BUTTON */}
        <section>
          <div className="dm-agent-controls">
            <input
              type="text"
              className="dm-agent-prompt-input"
              placeholder={`Optional Agent scraping instructions for "${activeProject.name}" -> "${activeTab?.name}" (e.g., specific company, cohort, or target URL)...`}
              value={agentPrompt}
              onChange={(e) => setAgentPrompt(e.target.value)}
            />

            <div
              className="dm-line-count-wrap"
              title="Request a specific number of new lines to produce (1 to 1000)"
            >
              <label htmlFor="dm-line-count-input" className="dm-line-count-label">
                Lines (max 1000):
              </label>
              <input
                id="dm-line-count-input"
                type="number"
                min={1}
                max={1000}
                className="dm-line-count-input"
                value={requestedLinesInput}
                onChange={(e) => {
                  const val = e.target.value
                  if (val === '') {
                    setRequestedLinesInput('')
                    return
                  }
                  const num = parseInt(val, 10)
                  if (!Number.isNaN(num)) {
                    setRequestedLinesInput(num > 1000 ? '1000' : val)
                  }
                }}
                onBlur={() => {
                  const num = parseInt(requestedLinesInput, 10)
                  setRequestedLinesInput(
                    String(Number.isNaN(num) ? 10 : Math.max(1, Math.min(1000, num)))
                  )
                }}
              />
            </div>

            <select
              className="dm-agent-select"
              value={agentMode}
              onChange={(e) =>
                setAgentMode(
                  e.target.value as
                    | 'gemini-3.8-flash'
                    | 'antigravity-agent'
                    | 'gemini-4-argon'
                )
              }
            >
              <option value="gemini-3.8-flash">Gemini 3.8 Flash (Web Search + Schema)</option>
              <option value="antigravity-agent">Antigravity Agent (Remote Sandbox)</option>
              <option value="gemini-4-argon">Gemini 4 Argon (Deep Reasoning)</option>
            </select>
          </div>

          <button
            type="button"
            className="dm-fetch-btn"
            disabled={fetchingData}
            onClick={handleFetchData}
          >
            {fetchingData
              ? `Agentic AI (${agentMode}) is producing ${
                  requestedLinesInput || 10
                } unique new lines formatted for "${activeTab?.name}"...`
              : `Fetch Data (${Math.max(
                  1,
                  Math.min(1000, parseInt(requestedLinesInput, 10) || 10)
                )} New ${
                  Math.max(1, Math.min(1000, parseInt(requestedLinesInput, 10) || 10)) === 1
                    ? 'Line'
                    : 'Lines'
                })`}
          </button>

          {fetchingData && (
            <div className="dm-agent-banner dm-fetch-progress">
              <strong>
                Searching… {fetchProgress?.found ?? 0} of{' '}
                {fetchProgress?.requested ?? requestedLinesInput} rows found
              </strong>
              {fetchProgress && fetchProgress.log.length > 0 ? (
                <ol className="dm-fetch-log">
                  {[...fetchProgress.log].reverse().slice(0, 12).map((entry, i) => (
                    <li key={`${entry.t}-${i}`}>
                      <span className="dm-fetch-log-time">{Math.round(entry.t)}s</span>
                      {entry.msg}
                    </li>
                  ))}
                </ol>
              ) : (
                <div className="dm-fetch-log-empty">Starting search…</div>
              )}
            </div>
          )}

          {!fetchingData && (newRows.length > 0 || activeProject.last_agent_run) && (
            <button
              type="button"
              className="dm-clear-new-data-btn"
              onClick={handleClearNewData}
            >
              Clear New Data
            </button>
          )}

          {!fetchingData && activeProject.last_agent_run && (
            <div className="dm-agent-banner">
              <strong>Agent Run ({activeProject.last_agent_run.model_used}):</strong>{' '}
              {activeProject.last_agent_run.agent_summary}
              {activeProject.last_agent_run.sources_consulted?.length > 0 && (
                <div style={{ marginTop: '4px', fontSize: '0.78rem', color: '#92400e' }}>
                  Sources: {activeProject.last_agent_run.sources_consulted.join(', ')}
                </div>
              )}
            </div>
          )}

          {/* 5. NEW DATA OUTPUT SECTION (ONLY NEW DATA, GOOGLE SHEETS FORMATTED & CLIPBOARD READY) */}
          {newRows.length > 0 && (
            <div className="dm-new-data-section">
              <div className="dm-section-label" style={{ marginTop: '24px', flexWrap: 'wrap', gap: '8px' }}>
                <div style={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
                  <strong>New Data Produced (Only New Data — No Duplicates)</strong>
                  <span
                    style={{
                      fontSize: '0.8rem',
                      color: '#065f46',
                      background: '#ecfdf5',
                      border: '1px solid #a7f3d0',
                      padding: '2px 8px',
                      borderRadius: '4px',
                      fontFamily: 'sans-serif',
                    }}
                  >
                    {newRows.length.toLocaleString()} new{' '}
                    {newRows.length === 1 ? 'line' : 'lines'} · Formatted like Google Sheets
                  </span>
                  {newRowsMergedStats.mergedGroupCount > 0 && (
                    <span className="dm-merged-detected-pill">
                      ⊞ {newRowsMergedStats.mergedGroupCount} Merged Row Groups (
                      {newRowsMergedStats.mergedRowCount} rows)
                    </span>
                  )}
                  {activeProject.columns.some((c) => c.put_link) && (
                    <span className="dm-merged-detected-pill" style={{ background: '#eff6ff', borderColor: '#bfdbfe', color: '#1d4ed8' }}>
                      🔗 Clickable Webpage Links Included
                    </span>
                  )}
                </div>
              </div>

              {copyToast && <div className="dm-copy-toast">{copyToast}</div>}

              <div className="dm-sheet-container">
                <div className="dm-sheet-toolbar">
                  <div className="dm-sheet-toolbar-left">
                    <strong>New Data Only</strong>
                    <span style={{ color: '#666' }}>
                      ({indexedFilteredNewRows.length.toLocaleString()}{' '}
                      {indexedFilteredNewRows.length === 1 ? 'row' : 'rows'})
                    </span>
                    <input
                      type="text"
                      className="dm-sheet-search"
                      placeholder="Filter new rows..."
                      value={newDataSearch}
                      onChange={(e) => {
                        setNewDataSearch(e.target.value)
                        setNewDataPage(1)
                      }}
                    />
                  </div>

                  <div className="dm-sheet-toolbar-right" style={{ flexWrap: 'wrap' }}>
                    <button
                      type="button"
                      className="dm-header-btn dm-copy-gsheet-btn"
                      title="Copies rich HTML table + TSV to clipboard (preserves clickable links, bullet points, and merged cells when pasted with Cmd+V in Google Sheets)"
                      onClick={() =>
                        handleCopyNewDataToClipboard(newRows, newDataHeaders, false, false)
                      }
                    >
                      📋 Copy for Google Sheets ({newRows.length.toLocaleString()} Rows)
                    </button>
                    <button
                      type="button"
                      className="dm-header-btn"
                      title="Copy with the column header row included"
                      onClick={() =>
                        handleCopyNewDataToClipboard(newRows, newDataHeaders, true, false)
                      }
                    >
                      📋 Copy with Headers
                    </button>
                    {activeProject.columns.some((c) => c.put_link) && (
                      <button
                        type="button"
                        className="dm-header-btn"
                        title="Copies plain-text TSV using =HYPERLINK(url, label) formulas for linked columns"
                        onClick={() =>
                          handleCopyNewDataToClipboard(newRows, newDataHeaders, false, true)
                        }
                      >
                        🔗 Copy with =HYPERLINK()
                      </button>
                    )}
                    <button
                      type="button"
                      className="dm-header-btn"
                      onClick={() => setShowRawTsvPreview((prev) => !prev)}
                    >
                      {showRawTsvPreview ? 'Hide Pasteable TSV' : 'Show Pasteable TSV'}
                    </button>
                    <button
                      type="button"
                      className="dm-header-btn"
                      onClick={() => handleExportNewDataCsv(newRows, newDataHeaders)}
                    >
                      Export New .CSV
                    </button>
                  </div>
                </div>

                {showRawTsvPreview && (
                  <div className="dm-tsv-preview-box">
                    <div className="dm-tsv-preview-header">
                      <span>
                        Tab-Separated Values (TSV) formatted for Google Sheets — multi-line bullet
                        cells are automatically wrapped in quotes so Google Sheets keeps all bullets
                        in a single cell. Use the green <strong>Copy for Google Sheets</strong>{' '}
                        button above to also include rich clickable links and merged cells:
                      </span>
                    </div>
                    <textarea
                      readOnly
                      className="dm-tsv-textarea"
                      value={buildGoogleSheetsTsv(
                        newRows,
                        newDataHeaders,
                        activeProject.columns,
                        false,
                        false
                      )}
                      onFocus={(e) => e.target.select()}
                    />
                  </div>
                )}

                <div className="dm-sheet-scroll">
                  <table className="dm-sheet-table">
                    <thead>
                      <tr>
                        <th className="dm-sheet-row-num">#</th>
                        {newDataHeaders.map((headerName) => (
                          <th key={headerName}>{headerName}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {visibleNewRows.map(({ row, originalIndex }, visIdx) => (
                        <tr key={originalIndex}>
                          <td className="dm-sheet-row-num">{originalIndex + 1}</td>
                          {newDataHeaders.map((headerName) => {
                            const cellKey = `${visIdx}:${headerName}`
                            if (visibleNewMergedLayout.hiddenSet.has(cellKey)) {
                              // Covered by a vertically merged `<td rowSpan>` on an earlier row
                              return null
                            }
                            const span = visibleNewMergedLayout.spanMap.get(cellKey) || 1
                            const matchingCol = activeProject.columns.find(
                              (c) =>
                                c.name.trim().toUpperCase() === headerName.trim().toUpperCase()
                            )
                            const cellVal = getCellValue(row, headerName)
                            const linkUrl = getCellLink(row, headerName, matchingCol)

                            return (
                              <MemoizedSheetCell
                                key={headerName}
                                rowIndex={originalIndex}
                                colName={headerName}
                                value={cellVal}
                                linkUrl={linkUrl}
                                span={span}
                                onCommit={handleCommitNewDataCell}
                              />
                            )
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {indexedFilteredNewRows.length > PAGE_SIZE && (
                  <div className="dm-sheet-pagination">
                    <div>
                      Showing {(safeNewPage - 1) * PAGE_SIZE + 1}–
                      {Math.min(safeNewPage * PAGE_SIZE, indexedFilteredNewRows.length)} of{' '}
                      {indexedFilteredNewRows.length.toLocaleString()} new rows (Copy button copies
                      all {newRows.length.toLocaleString()} new rows)
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <button
                        type="button"
                        className="dm-header-btn"
                        disabled={safeNewPage <= 1}
                        onClick={() => setNewDataPage(1)}
                      >
                        « First
                      </button>
                      <button
                        type="button"
                        className="dm-header-btn"
                        disabled={safeNewPage <= 1}
                        onClick={() => setNewDataPage((p) => Math.max(1, p - 1))}
                      >
                        ‹ Prev
                      </button>
                      <span style={{ padding: '0 8px' }}>
                        Page <strong>{safeNewPage}</strong> of <strong>{totalNewPages}</strong>
                      </span>
                      <button
                        type="button"
                        className="dm-header-btn"
                        disabled={safeNewPage >= totalNewPages}
                        onClick={() => setNewDataPage((p) => Math.min(totalNewPages, p + 1))}
                      >
                        Next ›
                      </button>
                      <button
                        type="button"
                        className="dm-header-btn"
                        disabled={safeNewPage >= totalNewPages}
                        onClick={() => setNewDataPage(totalNewPages)}
                      >
                        Last »
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}
        </section>
      </main>

      {/* Modal: Add Custom MCP Connector */}
      {showAddMcpModal && (
        <div className="dm-modal-backdrop" onClick={() => setShowAddMcpModal(false)}>
          <div className="dm-modal" onClick={(e) => e.stopPropagation()}>
            <h3>Add Custom MCP / Data Connector</h3>
            <p className="dm-auth-subtitle">
              Connect a remote Model Context Protocol (MCP) server or add a prioritized web source
              for the Gemini Agentic scraper.
            </p>
            <form onSubmit={handleCreateCustomMcp}>
              <div className="dm-auth-field">
                <label>Connector Name</label>
                <input
                  type="text"
                  required
                  placeholder="e.g., Crunchbase, GitHub, Substack"
                  value={newMcpName}
                  onChange={(e) => setNewMcpName(e.target.value)}
                />
              </div>
              <div className="dm-auth-field">
                <label>Remote MCP Server URL (Optional)</label>
                <input
                  type="text"
                  placeholder="https://mcp.example.com/sse"
                  value={newMcpUrl}
                  onChange={(e) => setNewMcpUrl(e.target.value)}
                />
              </div>
              <div style={{ display: 'flex', gap: '10px', justifyContent: 'flex-end' }}>
                <button
                  type="button"
                  className="dm-header-btn"
                  onClick={() => setShowAddMcpModal(false)}
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="dm-header-btn"
                  style={{ background: '#ffa000', borderColor: '#ffa000' }}
                >
                  Add MCP
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Modal: AI Settings */}
      {showAiKeyModal && (
        <div className="dm-modal-backdrop" onClick={() => setShowAiKeyModal(false)}>
          <div className="dm-modal" onClick={(e) => e.stopPropagation()}>
            <h3>Agentic AI Configuration</h3>
            <p className="dm-auth-subtitle">
              Powered by the Google GenAI Interactions API (`google-genai`). Uses{' '}
              <strong>Gemini 3.8 Flash</strong> with Google Search grounding, URL Context, and JSON
              Schema structured output.
            </p>
            <form onSubmit={handleSaveApiKey}>
              <div className="dm-auth-field">
                <label>Gemini API Key (optional if set in backend/.env)</label>
                <input
                  type="password"
                  placeholder="Configured in backend/.env"
                  value={geminiApiKey}
                  onChange={(e) => setGeminiApiKey(e.target.value)}
                />
              </div>
              <div className="dm-auth-field">
                <label>Backend API Base URL (optional override for Vercel / remote backend)</label>
                <input
                  type="text"
                  placeholder={
                    import.meta.env.VITE_API_BASE_URL || 'http://127.0.0.1:8000'
                  }
                  value={customApiBaseUrl}
                  onChange={(e) => setCustomApiBaseUrl(e.target.value)}
                />
              </div>
              <div style={{ display: 'flex', gap: '10px', justifyContent: 'flex-end' }}>
                <button
                  type="button"
                  className="dm-header-btn"
                  onClick={() => setShowAiKeyModal(false)}
                >
                  Close
                </button>
                <button
                  type="submit"
                  className="dm-header-btn"
                  style={{ background: '#ffa000', borderColor: '#ffa000' }}
                >
                  Save Settings
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  )
}

export default App
