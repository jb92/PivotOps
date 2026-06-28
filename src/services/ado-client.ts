/**
 * PivotOps â€” Azure DevOps REST API Client
 * Makes direct calls to ADO REST API from the browser using the user's token.
 * No server/proxy required â€” ADO supports CORS for authenticated requests.
 */

import { getAuthHeader } from "./auth";
import { storageGetItem, storageSetItem } from "./storage";

// â”€â”€ Types â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

export interface AdoConfig {
  organization: string;
  project: string;
}

export interface WorkItem {
  id: number;
  rev: number;
  fields: Record<string, unknown>;
  url: string;
}

export interface WorkItemUpdate {
  op: "add" | "replace" | "remove" | "test";
  path: string;
  value?: unknown;
}

export interface WiqlResult {
  workItems: Array<{ id: number; url: string }>;
}

export interface SyncOptions {
  layout: "flat" | "hierarchical";
  wiql?: string;
}

export interface TeamIteration {
  id: string;
  name: string;
  path: string;
  attributes: {
    startDate: string;
    finishDate: string;
    timeFrame: string;
  };
}

export interface AnalyticsWorkItem {
  WorkItemId: number;
  WorkItemType: string;
  Title: string;
  State: string;
  AssignedTo?: { UserName: string } | string;
  CreatedDate: string;
  ChangedDate: string;
  ClosedDate?: string;
  StoryPoints?: number;
  Iteration?: { IterationPath: string; };
  Area?: { AreaPath: string; };
  IterationPath?: string;
  AreaPath?: string;
  CycleTimeDays?: number;
  LeadTimeDays?: number;
  Priority?: number;
  Severity?: string;
  Tags?: string;
}

// â”€â”€ Config persistence â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

const ADO_CONFIG_KEY = "pivotops_ado_config";

export async function saveAdoConfig(config: AdoConfig): Promise<void> {
  await storageSetItem(ADO_CONFIG_KEY, JSON.stringify(config));
}

export async function getAdoConfig(): Promise<AdoConfig | null> {
  const raw = await storageGetItem(ADO_CONFIG_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw) as AdoConfig;
  } catch {
    return null;
  }
}

export async function hasAdoConfig(): Promise<boolean> {
  return !!(await getAdoConfig());
}

// â”€â”€ HTTP helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

async function adoFetch<T>(url: string, options: RequestInit = {}): Promise<T> {
  const authHeader = await getAuthHeader();

  // Only send Content-Type for requests with a body — sending it on GET requests
  // triggers a CORS preflight that some ADO endpoints may reject from localhost.
  const hasBody = options.method && ["POST", "PUT", "PATCH"].includes(options.method.toUpperCase());

  const response = await fetch(url, {
    ...options,
    mode: "cors",
    headers: {
      Authorization: authHeader,
      ...(hasBody ? { "Content-Type": "application/json" } : {}),
      Accept: "application/json",
      ...options.headers,
    },
  });

  if (!response.ok) {
    const body = await response.text();
    // Try to extract ADO's structured error message
    let detail = body;
    try {
      const json = JSON.parse(body) as { message?: string; typeKey?: string };
      if (json.message) detail = json.message;
    } catch { /* not JSON — use raw body */ }
    throw new Error(`ADO API error ${response.status}: ${detail}`);
  }

  return response.json() as Promise<T>;
}

function restUrl(org: string, project: string, path: string, apiVersion = "7.1"): string {
  return `https://dev.azure.com/${encodeURIComponent(org)}/${encodeURIComponent(project)}/_apis/${path}?api-version=${apiVersion}`;
}

function analyticsUrl(org: string, project: string, entitySet: string, query: string): string {
  return `https://analytics.dev.azure.com/${encodeURIComponent(org)}/${encodeURIComponent(project)}/_odata/v4.0-preview/${entitySet}?${query}`;
}

// â”€â”€ Read Operations â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

export async function queryWorkItems(wiql: string): Promise<WorkItem[]> {
  const config = await getAdoConfig();
  if (!config) throw new Error("Azure DevOps not configured.");

  const result = await adoFetch<WiqlResult>(
    restUrl(config.organization, config.project, "wit/wiql"),
    {
      method: "POST",
      body: JSON.stringify({ query: wiql }),
    }
  );

  if (result.workItems.length === 0) return [];

  // Batch fetch work items (max 200 per call)
  const ids = result.workItems.map((wi) => wi.id);
  return fetchWorkItemsByIds(ids);
}

export async function fetchWorkItemsByIds(ids: number[]): Promise<WorkItem[]> {
  const config = await getAdoConfig();
  if (!config) throw new Error("Azure DevOps not configured.");

  const allItems: WorkItem[] = [];
  const batchSize = 200;

  for (let i = 0; i < ids.length; i += batchSize) {
    const batch = ids.slice(i, i + batchSize);
    const idParam = batch.join(",");
    const url = `https://dev.azure.com/${encodeURIComponent(config.organization)}/_apis/wit/workitems?ids=${idParam}&api-version=7.1`;
    const result = await adoFetch<{ value: WorkItem[] }>(url);
    allItems.push(...result.value);
  }

  return allItems;
}

export async function getIterations(teamId?: string): Promise<TeamIteration[]> {
  const config = await getAdoConfig();
  if (!config) throw new Error("Azure DevOps not configured.");

  const team = teamId || `${config.project} Team`;
  const result = await adoFetch<{ value: TeamIteration[] }>(
    restUrl(
      config.organization,
      config.project,
      `work/teamsettings/iterations`
    ) + `&$team=${encodeURIComponent(team)}`
  );

  return result.value;
}

/**
 * Fetch allowed values for dropdown fields from all work item types in the project.
 * Returns a map: field refName â†’ string[] of allowed values.
 */
export async function fetchFieldAllowedValues(): Promise<Record<string, string[]>> {
  const config = await getAdoConfig();
  if (!config) throw new Error("Azure DevOps not configured.");

  const typesResult = await adoFetch<{ value: Array<{ name: string }> }>(
    restUrl(config.organization, config.project, "wit/workitemtypes")
  );

  const allowedMap: Record<string, Set<string>> = {};

  // Fetch field definitions for each work item type â€” collect all allowed values
  for (const wit of typesResult.value) {
    try {
      const detail = await adoFetch<{
        fields: Array<{ referenceName: string; allowedValues: string[] }>;
      }>(
        restUrl(config.organization, config.project, `wit/workitemtypes/${encodeURIComponent(wit.name)}`)
      );
      for (const field of detail.fields) {
        if (field.allowedValues && field.allowedValues.length > 0) {
          if (!allowedMap[field.referenceName]) {
            allowedMap[field.referenceName] = new Set();
          }
          for (const v of field.allowedValues) {
            allowedMap[field.referenceName].add(v);
          }
        }
      }
    } catch {
      // Skip types we can't read
    }
  }

  const result: Record<string, string[]> = {};
  for (const [key, valSet] of Object.entries(allowedMap)) {
    result[key] = [...valSet];
  }
  return result;
}

/**
 * Query work items with parentâ€“child hierarchy via a tree query.
 * Returns items with an extra _depth and _parentId field stuffed into fields.
 */
async function queryWorkItemsTree(wiql: string): Promise<WorkItem[]> {
  const config = await getAdoConfig();
  if (!config) throw new Error("Azure DevOps not configured.");

  // Use a tree query: SELECT ... FROM WorkItemLinks WHERE ...
  const treeWiql =
    "SELECT [System.Id], [System.Title], [System.State], [System.AssignedTo], " +
    "[System.WorkItemType], [Microsoft.VSTS.Scheduling.StoryPoints], " +
    "[System.IterationPath], [System.AreaPath], [System.Tags], " +
    "[System.CreatedDate], [System.ChangedDate] " +
    "FROM WorkItemLinks " +
    "WHERE ([Source].[System.TeamProject] = @project) " +
    "AND ([System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Forward') " +
    "AND ([Target].[System.TeamProject] = @project) " +
    "MODE (Recursive)";

  const result = await adoFetch<{
    workItemRelations: Array<{ target: { id: number }; source: { id: number } | null }>;
  }>(
    restUrl(config.organization, config.project, "wit/wiql"),
    { method: "POST", body: JSON.stringify({ query: wiql || treeWiql }) }
  );

  if (!result.workItemRelations || result.workItemRelations.length === 0) return [];

  // Build parent map and collect IDs
  const parentMap = new Map<number, number | null>();
  const ids: number[] = [];
  for (const rel of result.workItemRelations) {
    if (rel.target) {
      ids.push(rel.target.id);
      parentMap.set(rel.target.id, rel.source?.id ?? null);
    }
  }

  const uniqueIds = [...new Set(ids)];
  const items = await fetchWorkItemsByIds(uniqueIds);

  // Compute depth for each item
  const depthCache = new Map<number, number>();
  function getDepth(id: number): number {
    if (depthCache.has(id)) return depthCache.get(id)!;
    const parent = parentMap.get(id);
    const depth = parent == null ? 0 : getDepth(parent) + 1;
    depthCache.set(id, depth);
    return depth;
  }

  // Sort by relation order and annotate depth
  const itemMap = new Map(items.map((wi) => [wi.id, wi]));
  const ordered: WorkItem[] = [];
  for (const rel of result.workItemRelations) {
    if (rel.target) {
      const wi = itemMap.get(rel.target.id);
      if (wi) {
        // Stuff hierarchy info into fields for later use
        wi.fields["_pivotops_depth"] = getDepth(rel.target.id);
        wi.fields["_pivotops_parentId"] = rel.source?.id ?? "";
        ordered.push(wi);
        itemMap.delete(rel.target.id); // avoid duplicates
      }
    }
  }

  return ordered;
}

/**
 * Fetch work items via the Analytics OData endpoint.
 * This gives us CycleTimeDays, LeadTimeDays, and other analytics-enriched fields.
 */
export async function fetchAnalyticsWorkItems(
  filter?: string,
  top = 5000
): Promise<AnalyticsWorkItem[]> {
  const config = await getAdoConfig();
  if (!config) throw new Error("Azure DevOps not configured.");

  const select = [
    "WorkItemId", "WorkItemType", "Title", "State",
    "AssignedTo", "CreatedDate", "ChangedDate", "ClosedDate",
    "StoryPoints", "CycleTimeDays", "LeadTimeDays",
    "Priority", "Severity", "Tags",
  ].join(",");

  // Expand navigation properties for iteration, area, and assigned-to
  const expand = "Iteration($select=IterationPath),Area($select=AreaPath),AssignedTo($select=UserName)";

  let query = `$select=${select}&$expand=${expand}&$top=${top}&$orderby=ChangedDate desc`;
  if (filter) {
    query += `&$filter=${encodeURIComponent(filter)}`;
  }

  const result = await adoFetch<{ value: AnalyticsWorkItem[] }>(
    analyticsUrl(config.organization, config.project, "WorkItems", query)
  );

  // Flatten nested navigation properties so downstream code can use simple strings
  return result.value.map((item) => ({
    ...item,
    AssignedTo: typeof item.AssignedTo === "object" ? item.AssignedTo?.UserName || "" : item.AssignedTo ?? "",
    IterationPath: item.Iteration?.IterationPath ?? "",
    AreaPath: item.Area?.AreaPath ?? "",
  }));
}

/**
 * Fetch work item state transitions (for flow analytics).
 */
export async function fetchWorkItemRevisions(
  workItemId: number
): Promise<Array<{ id: number; rev: number; fields: Record<string, unknown> }>> {
  const config = await getAdoConfig();
  if (!config) throw new Error("Azure DevOps not configured.");

  const result = await adoFetch<{
    value: Array<{ id: number; rev: number; fields: Record<string, unknown> }>;
  }>(
    restUrl(config.organization, config.project, `wit/workitems/${workItemId}/revisions`)
  );

  return result.value;
}

// â”€â”€ Write Operations â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

export async function updateWorkItem(
  workItemId: number,
  updates: WorkItemUpdate[],
  retryCount = 0
): Promise<WorkItem> {
  const config = await getAdoConfig();
  if (!config) throw new Error("Azure DevOps not configured.");

  // Fetch current work item to filter out no-op changes and avoid 409 conflicts
  const current = (await fetchWorkItemsByIds([workItemId]))[0];
  if (!current) throw new Error(`Work item ${workItemId} not found.`);

  // Normalize an ADO field value to a comparable string
  function normalizeFieldValue(val: unknown): string {
    if (val === null || val === undefined) return "";
    if (typeof val === "object") {
      // ADO identity fields come back as objects with displayName/uniqueName
      const obj = val as Record<string, unknown>;
      return String(obj.displayName || obj.uniqueName || obj.name || "").trim();
    }
    // Normalize numbers and strings alike
    return String(val).trim().replace(/\s+/g, " ");
  }

  // HTML fields need special normalization (strip whitespace differences)
  const htmlFields = new Set(["System.Description"]);
  function normalizeHtml(val: string): string {
    return val.replace(/\r?\n/g, "").replace(/>\s+</g, "><").trim();
  }

  // Map patch paths to ADO field names
  const fieldMap: Record<string, string> = {
    "/fields/System.State": "System.State",
    "/fields/System.Title": "System.Title",
    "/fields/System.AssignedTo": "System.AssignedTo",
    "/fields/System.Description": "System.Description",
    "/fields/System.WorkItemType": "System.WorkItemType",
    "/fields/Microsoft.VSTS.Scheduling.StoryPoints": "Microsoft.VSTS.Scheduling.StoryPoints",
    "/fields/Microsoft.VSTS.Common.Priority": "Microsoft.VSTS.Common.Priority",
    "/fields/System.IterationPath": "System.IterationPath",
    "/fields/System.AreaPath": "System.AreaPath",
    "/fields/System.Tags": "System.Tags",
  };

  const actualChanges = updates.filter((u) => {
    const adoField = fieldMap[u.path];
    if (adoField) {
      let currentVal = normalizeFieldValue(current.fields[adoField]);
      let newVal = normalizeFieldValue(u.value);
      // HTML fields need extra normalization
      if (htmlFields.has(adoField)) {
        currentVal = normalizeHtml(currentVal);
        newVal = normalizeHtml(newVal);
      }
      return currentVal !== newVal;
    }
    return true; // Unknown fields — send them through
  });

  if (actualChanges.length === 0) {
    return current; // Nothing changed
  }

  try {
    return await adoFetch<WorkItem>(
      restUrl(config.organization, config.project, `wit/workitems/${workItemId}`),
      {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json-patch+json",
        },
        body: JSON.stringify(actualChanges),
      }
    );
  } catch (error) {
    // Retry on 409 (revision conflict) — re-fetch and try again
    if (retryCount < 2 && error instanceof Error && error.message.includes("409")) {
      return updateWorkItem(workItemId, updates, retryCount + 1);
    }
    throw error;
  }
}

export async function bulkUpdateWorkItems(
  updates: Array<{ id: number; changes: WorkItemUpdate[] }>
): Promise<{ succeeded: WorkItem[]; failed: Array<{ id: number; error: string }> }> {
  const succeeded: WorkItem[] = [];
  const failed: Array<{ id: number; error: string }> = [];

  for (const { id, changes } of updates) {
    try {
      const result = await updateWorkItem(id, changes);
      succeeded.push(result);
    } catch (error) {
      failed.push({
        id,
        error: error instanceof Error ? error.message : String(error),
      });
    }
  }

  return { succeeded, failed };
}

/**
 * Dry-run validation for bulk updates.
 * Checks that target fields/values are valid before calling the API.
 */
export async function validateUpdates(
  updates: Array<{ id: number; changes: WorkItemUpdate[] }>
): Promise<Array<{ id: number; warnings: string[] }>> {
  const config = await getAdoConfig();
  if (!config) throw new Error("Azure DevOps not configured.");

  // Fetch current state of all items to validate against
  const ids = updates.map((u) => u.id);
  const currentItems = await fetchWorkItemsByIds(ids);
  const itemMap = new Map(currentItems.map((wi) => [wi.id, wi]));

  const results: Array<{ id: number; warnings: string[] }> = [];

  for (const update of updates) {
    const warnings: string[] = [];
    const current = itemMap.get(update.id);

    if (!current) {
      warnings.push(`Work item ${update.id} not found.`);
      results.push({ id: update.id, warnings });
      continue;
    }

    for (const change of update.changes) {
      // Validate field path format
      if (!change.path.startsWith("/fields/")) {
        warnings.push(`Invalid field path: ${change.path}`);
      }

      // Check for empty values on 'replace' operations
      if (change.op === "replace" && (change.value === null || change.value === undefined)) {
        warnings.push(`Empty value for ${change.path}`);
      }
    }

    results.push({ id: update.id, warnings });
  }

  return results;
}

// â”€â”€ Data sync to Excel â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

// Color palette for the styled table
const COLORS = {
  headerBg: "#0078D4",
  headerFont: "#FFFFFF",
  rowEven: "#F3F6FC",
  rowOdd: "#FFFFFF",
  borderColor: "#D6DDE8",
  depthIndent: "#E8EEF5",
  idBg: "#EBF3FF",
  stateDone: "#DFF6DD",
  stateActive: "#FFF4CE",
  stateNew: "#EBF3FF",
  stateRemoved: "#FDE7E9",
  typeBug: "#F38BA8",
  typeEpic: "#CBA6F7",
  typeFeature: "#89B4FA",
  typeUserStory: "#A6E3A1",
  typeTask: "#FAB387",
};

const STATE_COLORS: Record<string, string> = {
  Done: COLORS.stateDone,
  Closed: COLORS.stateDone,
  Resolved: COLORS.stateDone,
  Active: COLORS.stateActive,
  "In Progress": COLORS.stateActive,
  Committed: COLORS.stateActive,
  New: COLORS.stateNew,
  Proposed: COLORS.stateNew,
  Removed: COLORS.stateRemoved,
};

export async function syncToExcel(options: SyncOptions = { layout: "flat" }): Promise<number> {
  const defaultQuery =
    "SELECT [System.Id], [System.Title], [System.State], [System.AssignedTo], " +
    "[System.WorkItemType], [Microsoft.VSTS.Scheduling.StoryPoints], " +
    "[System.IterationPath], [System.AreaPath], [System.Tags], " +
    "[System.CreatedDate], [System.ChangedDate] " +
    "FROM WorkItems WHERE [System.TeamProject] = @project ORDER BY [System.ChangedDate] DESC";

  // Fetch work items: flat or hierarchical
  const items = options.layout === "hierarchical"
    ? await queryWorkItemsTree(options.wiql ?? "")
    : await queryWorkItems(options.wiql || defaultQuery);

  // Fetch allowed values for dropdown columns
  let allowedValues: Record<string, string[]> = {};
  try {
    allowedValues = await fetchFieldAllowedValues();
  } catch {
    // Non-critical â€” proceed without dropdowns
  }

  await Excel.run(async (ctx) => {
    // Create or get the PivotOps data sheet
    let sheet = ctx.workbook.worksheets.getItemOrNullObject("PivotOps Data");
    await ctx.sync();
    if (sheet.isNullObject) {
      sheet = ctx.workbook.worksheets.add("PivotOps Data");
    } else {
      // Delete existing tables FIRST to avoid overlap errors
      const tables = sheet.tables;
      tables.load("items");
      await ctx.sync();
      for (const t of tables.items) {
        t.delete();
      }
      await ctx.sync();
      sheet.getRange().clear();
      await ctx.sync();
    }

    const isHierarchical = options.layout === "hierarchical";

    // â”€â”€ Build columns dynamically from actual work item fields â”€â”€
    // Known field reference names â†’ friendly header names
    const knownHeaders: Record<string, string> = {
      "System.Id": "ID",
      "System.WorkItemType": "Type",
      "System.Title": "Title",
      "System.State": "State",
      "System.AssignedTo": "Assigned To",
      "Microsoft.VSTS.Scheduling.StoryPoints": "Story Points",
      "System.IterationPath": "Iteration",
      "System.AreaPath": "Area Path",
      "System.Tags": "Tags",
      "System.CreatedDate": "Created",
      "System.ChangedDate": "Changed",
      "Microsoft.VSTS.Common.ValueArea": "Value Area",
      "Microsoft.VSTS.Common.Priority": "Priority",
      "Microsoft.VSTS.Common.Severity": "Severity",
      "Microsoft.VSTS.Scheduling.RemainingWork": "Remaining Work",
      "Microsoft.VSTS.Scheduling.OriginalEstimate": "Original Estimate",
      "Microsoft.VSTS.Scheduling.CompletedWork": "Completed Work",
      "Microsoft.VSTS.Common.Activity": "Activity",
      "System.Reason": "Reason",
      "System.Description": "Description",
      "System.Parent": "Parent",
    };

    // Discover all field ref names present in the fetched items
    const fieldRefNames = new Set<string>();
    // Always include System.Id first
    fieldRefNames.add("System.Id");
    for (const item of items) {
      for (const key of Object.keys(item.fields)) {
        if (!key.startsWith("_pivotops_")) {
          fieldRefNames.add(key);
        }
      }
    }

    // Default field order (for sorting known fields to the front)
    const defaultOrder = [
      "System.Id", "System.WorkItemType", "System.Title", "System.State",
      "System.AssignedTo", "Microsoft.VSTS.Scheduling.StoryPoints",
      "System.IterationPath", "System.AreaPath", "System.Tags",
      "System.CreatedDate", "System.ChangedDate",
    ];

    // Sort: known fields in default order first, then remaining fields alphabetically
    const fieldMap = [...fieldRefNames].sort((a, b) => {
      const ia = defaultOrder.indexOf(a);
      const ib = defaultOrder.indexOf(b);
      if (ia >= 0 && ib >= 0) return ia - ib;
      if (ia >= 0) return -1;
      if (ib >= 0) return 1;
      return a.localeCompare(b);
    });

    // Build header names
    const dataHeaders = fieldMap.map((ref) => {
      if (knownHeaders[ref]) return knownHeaders[ref];
      // Extract last segment: "Microsoft.VSTS.Common.ValueArea" â†’ "ValueArea" â†’ "Value Area"
      const last = ref.split(".").pop() || ref;
      return last.replace(/([a-z])([A-Z])/g, "$1 $2");
    });

    const headers = isHierarchical ? ["Level", ...dataHeaders] : dataHeaders;

    // Fields that should get dropdown data-validation
    const dropdownFields: Record<string, string> = {
      "System.State": "State",
      "System.WorkItemType": "Type",
    };

    // â”€â”€ Write headers â”€â”€
    const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
    headerRange.values = [headers];
    headerRange.format.font.bold = true;
    headerRange.format.font.color = COLORS.headerFont;
    headerRange.format.font.size = 12;
    headerRange.format.fill.color = COLORS.headerBg;
    headerRange.format.rowHeight = 32;
    headerRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;
    headerRange.format.verticalAlignment = Excel.VerticalAlignment.center;

    if (items.length === 0) {
      sheet.getUsedRange().format.autofitColumns();
      sheet.activate();
      await ctx.sync();
      return;
    }

    // â”€â”€ Build data rows â”€â”€
    const colOffset = isHierarchical ? 1 : 0;

    const data = items.map((item) => {
      const depth = isHierarchical ? (item.fields["_pivotops_depth"] as number ?? 0) : 0;
      const indent = isHierarchical && depth > 0 ? "  ".repeat(depth) + "> " : "";

      const row = fieldMap.map((field) => {
        // System.Id lives at the top level of the work item, not inside fields
        const val = field === "System.Id" ? item.id : item.fields[field];
        // Format assigned-to as display name
        if (val && typeof val === "object" && "displayName" in (val as Record<string, unknown>)) {
          return (val as { displayName: string }).displayName;
        }
        // For Title column in hierarchical mode, add indent prefix
        if (field === "System.Title" && isHierarchical) {
          return indent + (val ?? "");
        }
        // Format date fields nicely (any field ending in Date)
        if (field.endsWith("Date") && val && typeof val === "string") {
          try {
            return new Date(val).toLocaleDateString("en-US", { year: "numeric", month: "short", day: "numeric" });
          } catch { return val; }
        }
        return val ?? "";
      });

      if (isHierarchical) {
        // Prepend depth indicator column
        const depthLabel = depth === 0 ? "Epic/Feature" : depth === 1 ? "Story/PBI" : "Task/Sub";
        return [depthLabel, ...row];
      }
      return row;
    });

    // Write all data at once
    const fullRange = sheet.getRangeByIndexes(0, 0, data.length + 1, headers.length);
    fullRange.values = [headers, ...data];

    // â”€â”€ Create Excel table â”€â”€
    const tableRange = sheet.getRangeByIndexes(0, 0, data.length + 1, headers.length);
    const table = sheet.tables.add(tableRange, true);
    table.name = "PivotOpsData";
    table.style = "TableStyleMedium2";
    table.showBandedRows = true;
    table.showFilterButton = true;

    // â”€â”€ Style data rows â”€â”€
    for (let r = 0; r < data.length; r++) {
      const rowRange = sheet.getRangeByIndexes(r + 1, 0, 1, headers.length);
      rowRange.format.rowHeight = 26;
      rowRange.format.verticalAlignment = Excel.VerticalAlignment.center;

      // Alternating row colors
      const bgColor = r % 2 === 0 ? COLORS.rowEven : COLORS.rowOdd;
      rowRange.format.fill.color = bgColor;

      // â”€â”€ Color-code State column â”€â”€
      const stateColIdx = headers.indexOf("State");
      if (stateColIdx >= 0) {
        const stateVal = String(data[r][stateColIdx]);
        const stateColor = STATE_COLORS[stateVal];
        if (stateColor) {
          const stateCell = sheet.getRangeByIndexes(r + 1, stateColIdx, 1, 1);
          stateCell.format.fill.color = stateColor;
          stateCell.format.font.bold = true;
        }
      }

      // â”€â”€ Color-code Type column â”€â”€
      const typeColIdx = headers.indexOf("Type");
      if (typeColIdx >= 0) {
        const typeVal = String(data[r][typeColIdx]).toLowerCase();
        const typeCell = sheet.getRangeByIndexes(r + 1, typeColIdx, 1, 1);
        if (typeVal.includes("bug")) typeCell.format.font.color = COLORS.typeBug;
        else if (typeVal.includes("epic")) typeCell.format.font.color = COLORS.typeEpic;
        else if (typeVal.includes("feature")) typeCell.format.font.color = COLORS.typeFeature;
        else if (typeVal.includes("user story") || typeVal.includes("product backlog")) typeCell.format.font.color = COLORS.typeUserStory;
        else if (typeVal.includes("task")) typeCell.format.font.color = COLORS.typeTask;
        typeCell.format.font.bold = true;
      }

      // â”€â”€ Style ID column â”€â”€
      const idColIdx = headers.indexOf("ID");
      if (idColIdx >= 0) {
        const idCell = sheet.getRangeByIndexes(r + 1, idColIdx, 1, 1);
        idCell.format.fill.color = COLORS.idBg;
        idCell.format.font.bold = true;
        idCell.format.horizontalAlignment = Excel.HorizontalAlignment.center;
      }

      // â”€â”€ Hierarchical depth indentation styling â”€â”€
      if (isHierarchical) {
        const depth = items[r].fields["_pivotops_depth"] as number ?? 0;
        if (depth > 0) {
          const levelCell = sheet.getRangeByIndexes(r + 1, 0, 1, 1);
          levelCell.format.font.size = 10;
          levelCell.format.font.color = "#6c7086";
          levelCell.format.font.italic = true;
        }
      }
    }

    // â”€â”€ Column widths for readability â”€â”€
    const colWidths: Record<string, number> = {
      "Level": 90, "ID": 60, "Type": 110, "Title": 350, "State": 100,
      "Assigned To": 150, "Story Points": 85, "Iteration": 180,
      "Area Path": 180, "Tags": 150, "Created": 110, "Changed": 110,
      "Value Area": 110, "Priority": 75, "Severity": 90, "Activity": 100,
      "Remaining Work": 100, "Original Estimate": 110, "Completed Work": 100,
      "Reason": 100, "Parent": 70,
    };
    for (let c = 0; c < headers.length; c++) {
      const w = colWidths[headers[c]];
      if (w) {
        sheet.getRangeByIndexes(0, c, 1, 1).format.columnWidth = w;
      } else {
        // Default width for unknown columns
        sheet.getRangeByIndexes(0, c, 1, 1).format.columnWidth = 120;
      }
    }

    // â”€â”€ Freeze header row â”€â”€
    sheet.freezePanes.freezeRows(1);

    // â”€â”€ Story Points column: number format + right-align â”€â”€
    const spColIdx = headers.indexOf("Story Points");
    if (spColIdx >= 0 && data.length > 0) {
      const spRange = sheet.getRangeByIndexes(1, spColIdx, data.length, 1);
      spRange.numberFormat = [["0"]];
      spRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;
    }

    // â”€â”€ Add dropdown data validation for State and Type â”€â”€
    for (const [fieldRef, headerName] of Object.entries(dropdownFields)) {
      const colIdx = headers.indexOf(headerName);
      if (colIdx < 0) continue;
      const values = allowedValues[fieldRef];
      if (!values || values.length === 0) continue;

      const validationRange = sheet.getRangeByIndexes(1, colIdx, data.length, 1);
      validationRange.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: values.join(","),
        },
      };
    }

    // â”€â”€ Add dropdown for Priority if present â”€â”€
    if (allowedValues["Microsoft.VSTS.Common.Priority"]) {
      // Priority isn't in our current field list, but if the user adds it this is ready
    }

    // â”€â”€ Add iteration dropdown from fetched iterations â”€â”€
    const iterColIdx = headers.indexOf("Iteration");
    if (iterColIdx >= 0 && data.length > 0) {
      // Collect unique iteration values from the data
      const uniqueIters = [...new Set(data.map(row => String(row[iterColIdx])).filter(Boolean))];
      if (uniqueIters.length > 0 && uniqueIters.length < 256) {
        const iterRange = sheet.getRangeByIndexes(1, iterColIdx, data.length, 1);
        iterRange.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: uniqueIters.join(","),
          },
        };
      }
    }

    sheet.activate();
    await ctx.sync();
  });

  return items.length;
}
