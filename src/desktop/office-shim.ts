/**
 * PivotOps — Office.js compatibility shim for the WinUI/WebView2 shell.
 *
 * Implements the subset of `Office`, `OfficeRuntime` and `Excel` that the task
 * pane uses, forwarding workbook operations to the native host, which applies
 * them to a real .xlsx file. Importing this module installs the globals, so it
 * must be evaluated before any add-in code runs.
 */

import { invoke, isDesktopHost } from "./bridge";

type Op = Record<string, unknown>;
type ResultHandler = (results: Record<string, unknown>) => void;

let keySeq = 0;
function nextKey(): string {
  return `k${++keySeq}`;
}

class RequestContext {
  readonly workbook: Workbook;
  private ops: Op[] = [];
  private handlers: ResultHandler[] = [];

  constructor() {
    this.workbook = new Workbook(this);
  }

  /** Queues an operation and returns it so callers can keep mutating it before the flush. */
  enqueue<T extends Op>(op: T): T {
    this.ops.push(op);
    return op;
  }

  expect(handler: ResultHandler): void {
    this.handlers.push(handler);
  }

  async sync(): Promise<void> {
    if (this.ops.length === 0 && this.handlers.length === 0) return;
    const ops = this.ops;
    const handlers = this.handlers;
    this.ops = [];
    this.handlers = [];
    const response = await invoke<{ results?: Record<string, unknown> }>("excel.sync", { ops });
    const results = response?.results ?? {};
    for (const handler of handlers) handler(results);
  }
}

class Workbook {
  readonly worksheets: WorksheetCollection;

  constructor(ctx: RequestContext) {
    this.worksheets = new WorksheetCollection(ctx);
  }
}

class WorksheetCollection {
  constructor(private readonly ctx: RequestContext) {}

  getItem(name: string): Worksheet {
    return new Worksheet(this.ctx, name);
  }

  getItemOrNullObject(name: string): Worksheet {
    const sheet = new Worksheet(this.ctx, name);
    const key = nextKey();
    this.ctx.enqueue({ t: "sheetGetOrNull", sheet: name, key });
    this.ctx.expect((results) => {
      if (key in results) sheet.isNullObject = results[key] === true;
    });
    return sheet;
  }

  add(name: string): Worksheet {
    this.ctx.enqueue({ t: "sheetAdd", sheet: name });
    return new Worksheet(this.ctx, name);
  }
}

class Worksheet {
  isNullObject = false;
  readonly tables: TableCollection;
  readonly freezePanes: FreezePanes;

  constructor(private readonly ctx: RequestContext, readonly name: string) {
    this.tables = new TableCollection(ctx, name);
    this.freezePanes = new FreezePanes(ctx, name);
  }

  /** Without an address this targets the whole sheet, matching Office.js. */
  getRange(): Range {
    return new Range(this.ctx, this.name, null);
  }

  getRangeByIndexes(rowIndex: number, columnIndex: number, rowCount: number, columnCount: number): Range {
    return new Range(this.ctx, this.name, { r: rowIndex, c: columnIndex, rows: rowCount, cols: columnCount });
  }

  getUsedRange(): Range {
    return new Range(this.ctx, this.name, null);
  }

  activate(): void {
    this.ctx.enqueue({ t: "activate", sheet: this.name });
  }
}

interface Area {
  r: number;
  c: number;
  rows: number;
  cols: number;
}

class TableCollection {
  items: Table[] = [];

  constructor(private readonly ctx: RequestContext, private readonly sheet: string) {}

  load(_properties?: string | string[]): TableCollection {
    const key = nextKey();
    this.ctx.enqueue({ t: "tablesLoad", sheet: this.sheet, key });
    this.ctx.expect((results) => {
      const names = (results[key] as string[]) ?? [];
      this.items = names.map((name) => new Table(this.ctx, this.sheet, null, name));
    });
    return this;
  }

  add(range: Range, hasHeaders: boolean): Table {
    const id = nextKey();
    const area = range.area;
    this.ctx.enqueue({
      t: "tableAdd",
      sheet: this.sheet,
      id,
      r: area?.r ?? 0,
      c: area?.c ?? 0,
      rows: area?.rows ?? 0,
      cols: area?.cols ?? 0,
      hasHeaders,
    });
    return new Table(this.ctx, this.sheet, id, null);
  }
}

class Table {
  constructor(
    private readonly ctx: RequestContext,
    private readonly sheet: string,
    private readonly id: string | null,
    private readonly existingName: string | null,
  ) {}

  private set(prop: string, value: unknown): void {
    this.ctx.enqueue({ t: "tableSet", sheet: this.sheet, id: this.id, target: this.existingName, prop, value });
  }

  set name(value: string) {
    this.set("name", value);
  }
  set style(value: string) {
    this.set("style", value);
  }
  set showBandedRows(value: boolean) {
    this.set("showBandedRows", value);
  }
  set showFilterButton(value: boolean) {
    this.set("showFilterButton", value);
  }

  delete(): void {
    this.ctx.enqueue({ t: "tableDelete", sheet: this.sheet, id: this.id, target: this.existingName });
  }
}

class FreezePanes {
  constructor(private readonly ctx: RequestContext, private readonly sheet: string) {}

  freezeRows(count: number): void {
    this.ctx.enqueue({ t: "freezeRows", sheet: this.sheet, count });
  }
}

class Range {
  readonly format: RangeFormat;
  readonly dataValidation: DataValidation;
  private loaded: unknown[][] | null = null;

  constructor(
    private readonly ctx: RequestContext,
    private readonly sheet: string,
    readonly area: Area | null,
  ) {
    this.format = new RangeFormat(ctx, sheet, area);
    this.dataValidation = new DataValidation(ctx, sheet, area);
  }

  set values(values: unknown[][]) {
    this.ctx.enqueue({ t: "setValues", sheet: this.sheet, ...this.coords(), values });
  }

  get values(): unknown[][] {
    return this.loaded as unknown[][];
  }

  set numberFormat(formats: unknown[][]) {
    this.ctx.enqueue({ t: "numberFormat", sheet: this.sheet, ...this.coords(), formats });
  }

  clear(): void {
    this.ctx.enqueue({ t: "clear", sheet: this.sheet, ...this.coords() });
  }

  load(_properties?: string | string[]): Range {
    const key = nextKey();
    this.ctx.enqueue({ t: "getValues", sheet: this.sheet, ...this.coords(), key });
    this.ctx.expect((results) => {
      this.loaded = (results[key] as unknown[][]) ?? null;
    });
    return this;
  }

  private coords(): Partial<Area> {
    return this.area ?? {};
  }
}

class RangeFormat {
  private op: Op | null = null;
  readonly font: FontFormat;
  readonly fill: FillFormat;

  constructor(
    private readonly ctx: RequestContext,
    private readonly sheet: string,
    private readonly area: Area | null,
  ) {
    this.font = new FontFormat(this);
    this.fill = new FillFormat(this);
  }

  /**
   * Collapses every property assignment on this range into a single queued op,
   * preserving the position at which the first assignment happened.
   */
  setProp(name: string, value: unknown): void {
    if (!this.op) {
      this.op = this.ctx.enqueue({
        t: "format",
        sheet: this.sheet,
        ...(this.area ?? {}),
        props: {} as Record<string, unknown>,
      });
    }
    (this.op.props as Record<string, unknown>)[name] = value;
  }

  set rowHeight(value: number) {
    this.setProp("rowHeight", value);
  }
  set columnWidth(value: number) {
    this.setProp("columnWidth", value);
  }
  set horizontalAlignment(value: string) {
    this.setProp("horizontalAlignment", value);
  }
  set verticalAlignment(value: string) {
    this.setProp("verticalAlignment", value);
  }

  autofitColumns(): void {
    this.ctx.enqueue({ t: "autofit", sheet: this.sheet, ...(this.area ?? {}) });
  }
}

class FontFormat {
  constructor(private readonly owner: RangeFormat) {}
  set bold(value: boolean) {
    this.owner.setProp("fontBold", value);
  }
  set italic(value: boolean) {
    this.owner.setProp("fontItalic", value);
  }
  set color(value: string) {
    this.owner.setProp("fontColor", value);
  }
  set size(value: number) {
    this.owner.setProp("fontSize", value);
  }
  set name(value: string) {
    this.owner.setProp("fontName", value);
  }
}

class FillFormat {
  constructor(private readonly owner: RangeFormat) {}
  set color(value: string) {
    this.owner.setProp("fillColor", value);
  }
}

class DataValidation {
  constructor(
    private readonly ctx: RequestContext,
    private readonly sheet: string,
    private readonly area: Area | null,
  ) {}

  set rule(value: { list?: { inCellDropDown?: boolean; source?: unknown } }) {
    const source = value?.list?.source;
    if (typeof source !== "string") return;
    this.ctx.enqueue({
      t: "validationList",
      sheet: this.sheet,
      ...(this.area ?? {}),
      source,
      inCellDropDown: value.list?.inCellDropDown !== false,
    });
  }
}

// ── Global installation ────────────────────────────────────────────────────

const ExcelShim = {
  async run<T>(callback: (ctx: RequestContext) => Promise<T>): Promise<T> {
    const ctx = new RequestContext();
    const result = await callback(ctx);
    await ctx.sync();
    return result;
  },
  HorizontalAlignment: {
    general: "General",
    left: "Left",
    center: "Center",
    right: "Right",
    fill: "Fill",
    justify: "Justify",
    centerAcrossSelection: "CenterAcrossSelection",
    distributed: "Distributed",
  },
  VerticalAlignment: {
    top: "Top",
    center: "Center",
    bottom: "Bottom",
    justify: "Justify",
    distributed: "Distributed",
  },
};

function domReady(): Promise<void> {
  if (document.readyState === "loading") {
    return new Promise((resolve) => document.addEventListener("DOMContentLoaded", () => resolve(), { once: true }));
  }
  return Promise.resolve();
}

const readyInfo = { host: "Excel", platform: "PC" };

const OfficeShim = {
  HostType: { Excel: "Excel" },
  PlatformType: { PC: "PC" },
  context: { host: "Excel", platform: "PC" },
  onReady(callback?: (info: typeof readyInfo) => unknown): Promise<typeof readyInfo> {
    const ready = domReady().then(() => readyInfo);
    if (callback) return ready.then((info) => Promise.resolve(callback(info)).then(() => info));
    return ready;
  },
  actions: { associate: () => undefined },
};

const globals = globalThis as unknown as Record<string, unknown>;
globals.Excel = ExcelShim;
globals.Office = OfficeShim;

// Without a host, storage.ts falls back to localStorage on its own.
if (isDesktopHost()) {
  globals.OfficeRuntime = {
    storage: {
      getItem: (key: string) => invoke<string | null>("storage.get", { key }),
      setItem: (key: string, value: string) => invoke<void>("storage.set", { key, value }),
      removeItem: (key: string) => invoke<void>("storage.remove", { key }),
    },
  };
}
