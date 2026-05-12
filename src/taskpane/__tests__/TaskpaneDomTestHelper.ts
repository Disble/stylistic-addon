import type { ReadyCallback } from "./TaskpaneTestHelper.types";
import { vi } from "vitest";

/**
 * Escapes user-provided text before fake DOM HTML snapshots.
 */
function escapeHtml(value: string): string {
  return value
    .split("&")
    .join("&amp;")
    .split("<")
    .join("&lt;")
    .split(">")
    .join("&gt;")
    .split('"')
    .join("&quot;")
    .split("'")
    .join("&#39;");
}

/** Converts data-* attribute names to the fake DOM dataset camelCase key. */
function toDatasetKey(value: string): string {
  const [head = "", ...tail] = value.split("-");
  return tail.reduce(
    (datasetKey, segment) =>
      datasetKey + (segment.length === 0 ? "" : segment[0].toUpperCase() + segment.slice(1)),
    head
  );
}

/** Minimal DOMTokenList replacement used by taskpane tests. */
export class FakeClassList {
  private readonly classes = new Set<string>();

  add(...names: string[]) {
    for (const name of names) {
      this.classes.add(name);
    }
  }

  remove(...names: string[]) {
    for (const name of names) {
      this.classes.delete(name);
    }
  }

  toggle(name: string): boolean {
    if (this.classes.has(name)) {
      this.classes.delete(name);
      return false;
    }

    this.classes.add(name);
    return true;
  }

  contains(name: string): boolean {
    return this.classes.has(name);
  }

  toString(): string {
    return Array.from(this.classes).join(" ");
  }
}

/** Fake DOM element used to exercise taskpane flows without a browser. */
export class FakeElement {
  style = { display: "", width: "" };
  dataset: Record<string, string> = {};

  get className(): string {
    return this.classList.toString();
  }

  set className(value: string) {
    const classList = new FakeClassList();
    for (const name of value.split(/\s+/).filter(Boolean)) {
      classList.add(name);
    }
    this._classList = classList;
  }

  private _classList = new FakeClassList();

  get classList(): FakeClassList {
    return this._classList;
  }

  disabled = false;
  value = "";
  onclick: ((ev: MouseEvent) => any) | null = null;
  children: FakeElement[] = [];
  parentElement: FakeElement | null = null;

  private text = "";
  private html = "";
  private readonly listeners = new Map<string, Array<(ev: any) => void>>();

  get textContent(): string {
    return this.text;
  }

  set textContent(value: string) {
    this.text = value ?? "";
    this.html = escapeHtml(this.text);
  }

  get innerHTML(): string {
    return this.html;
  }

  set innerHTML(value: string) {
    this.html = value;
    if (value === "") {
      for (const child of this.children) {
        child.parentElement = null;
      }
      this.children = [];
    }
  }

  appendChild(child: FakeElement): FakeElement {
    if (child.parentElement) {
      const previousIndex = child.parentElement.children.indexOf(child);
      if (previousIndex !== -1) {
        child.parentElement.children.splice(previousIndex, 1);
      }
    }

    child.parentElement = this;
    this.children.push(child);
    return child;
  }

  /** Inserts a child before the given reference node, matching DOM semantics. */
  insertBefore(child: FakeElement, referenceNode: FakeElement | null): FakeElement {
    if (referenceNode === null) {
      return this.appendChild(child);
    }

    if (child.parentElement) {
      const previousIndex = child.parentElement.children.indexOf(child);
      if (previousIndex !== -1) {
        child.parentElement.children.splice(previousIndex, 1);
      }
    }

    const referenceIndex = this.children.indexOf(referenceNode);
    if (referenceIndex === -1) {
      throw new Error("Reference node is not a child of this element");
    }

    child.parentElement = this;
    this.children.splice(referenceIndex, 0, child);
    return child;
  }

  /** Inserts this node immediately before a sibling, matching Element.before(). */
  before(node: FakeElement): void {
    if (!this.parentElement) {
      return;
    }

    this.parentElement.insertBefore(node, this);
  }

  addEventListener(event: string, handler: (ev: any) => void) {
    if (!this.listeners.has(event)) {
      this.listeners.set(event, []);
    }
    this.listeners.get(event)?.push(handler);
  }

  click() {
    if (this.onclick) {
      this.onclick({} as MouseEvent);
    }
    for (const handler of this.listeners.get("click") ?? []) {
      handler({});
    }
  }

  querySelector(selector: string): FakeElement | null {
    for (const child of this.children) {
      if (matchesSelector(child, selector)) {
        return child;
      }
      const nested = child.querySelector(selector);
      if (nested) {
        return nested;
      }
    }
    return null;
  }

  querySelectorAll(selector: string): FakeElement[] {
    const results: FakeElement[] = [];
    for (const child of this.children) {
      if (matchesSelector(child, selector)) {
        results.push(child);
      }
      results.push(...child.querySelectorAll(selector));
    }
    return results;
  }

  remove() {
    if (this.parentElement) {
      const index = this.parentElement.children.indexOf(this);
      if (index !== -1) {
        this.parentElement.children.splice(index, 1);
      }
      this.parentElement = null;
    }
  }

  setAttribute(name: string, value: string) {
    (this as any)[`_attr_${name}`] = value;
    if (name.startsWith("data-")) {
      const datasetKey = toDatasetKey(name.slice(5));
      this.dataset[datasetKey] = value;
    }
  }

  getAttribute(name: string): string | null {
    return (this as any)[`_attr_${name}`] ?? null;
  }
}

function matchesSelector(el: FakeElement, selector: string): boolean {
  const attrEqMatch = /^\[([^\]="]+)="([^"]+)"\]$/.exec(selector);
  if (attrEqMatch) {
    const [, attr, value] = attrEqMatch;
    if (attr.startsWith("data-")) {
      const datasetKey = toDatasetKey(attr.slice(5));
      return el.dataset[datasetKey] === value;
    }

    return (el as any)[attr] === value || (el as any)[`_attr_${attr}`] === value;
  }

  const attrMatch = /^\[([^\]="]+)\]$/.exec(selector);
  if (attrMatch) {
    const [, attr] = attrMatch;
    if (attr.startsWith("data-")) {
      const datasetKey = toDatasetKey(attr.slice(5));
      return el.dataset[datasetKey] !== undefined;
    }

    return (el as any)[attr] !== undefined || (el as any)[`_attr_${attr}`] !== undefined;
  }

  const classMatch = /^\.(.+)$/.exec(selector);
  if (classMatch) {
    return el.classList.contains(classMatch[1]);
  }

  return false;
}

/** Fake document implementation used by the taskpane harness. */
export class FakeDocument {
  private readonly elements = new Map<string, FakeElement>();

  constructor(ids: string[]) {
    for (const id of ids) {
      this.elements.set(id, new FakeElement());
    }
  }

  getElementById(id: string): FakeElement | null {
    return this.elements.get(id) ?? null;
  }

  createElement(_tagName?: string): FakeElement {
    return new FakeElement();
  }
}

/** Returns a required fake element by id or throws in tests. */
function getRequiredElement(doc: FakeDocument, id: string): FakeElement {
  const el = doc.getElementById(id);
  if (!el) {
    throw new Error(`Missing fake DOM element: ${id}`);
  }
  return el;
}

/** Creates the fake taskpane DOM with the default initial visibility state. */
export function createTaskpaneDocument(): FakeDocument {
  const doc = new FakeDocument([
    "app-body",
    "container",
    "btn-analyze",
    "btn-analyze-label",
    "btn-cleanup",
    "btn-cleanup-label",
    "profile-select",
    "status-bar",
    "progress-container",
    "progress-bar",
    "progress-text",
    "results-panel",
    "results-summary",
    "results-list",
    "cleanup-section",
    "disable-track-changes-section",
    "btn-disable-track-changes",
    "btn-disable-track-changes-label",
  ]);

  for (const [elementId, display] of [
    ["app-body", "flex"],
    ["cleanup-section", "none"],
    ["disable-track-changes-section", "none"],
    ["results-panel", "block"],
    ["progress-container", "none"],
  ] as const) {
    getRequiredElement(doc, elementId).style.display = display;
  }

  getRequiredElement(doc, "profile-select").value = "narrativa-literaria";
  getRequiredElement(doc, "btn-analyze-label").textContent = "Analizar y sugerir";
  getRequiredElement(doc, "btn-cleanup-label").textContent = "Limpiar comentarios resueltos";
  getRequiredElement(doc, "btn-disable-track-changes-label").textContent =
    "Desactivar control de cambios";

  return doc;
}

/** Creates a controllable `Office.onReady` harness for taskpane tests. */
export function createOffice() {
  let readyCallback: ReadyCallback | undefined;
  const office = {
    HostType: { Word: "Word" },
    onReady: vi.fn((callback: ReadyCallback) => {
      readyCallback = callback;
      return Promise.resolve();
    }),
  };

  return {
    office,
    triggerReady(info: { host: string }) {
      if (!readyCallback) {
        throw new Error("Office.onReady callback was not registered");
      }
      readyCallback(info);
    },
  };
}
