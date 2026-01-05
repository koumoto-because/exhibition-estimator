// src/renderer/utils/jsonPatch.ts
// RFC6902 JSON Patch generator: deep path, 1 op per leaf change
// - object: recurse key-by-key (add/remove/replace)
// - array: recurse index-by-index + add/remove tail
// - primitive: replace
// JSON Pointer escaping: "~" -> "~0", "/" -> "~1"

export type JsonPatchOp =
  | { op: "add"; path: string; value: any }
  | { op: "remove"; path: string }
  | { op: "replace"; path: string; value: any };

type AnyJson = any;

function isObject(v: AnyJson): v is Record<string, AnyJson> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

function escapeJsonPointerToken(token: string): string {
  return token.replace(/~/g, "~0").replace(/\//g, "~1");
}

function joinPath(base: string, token: string): string {
  const t = escapeJsonPointerToken(token);
  return base === "" ? `/${t}` : `${base}/${t}`;
}

function valuesEqual(a: AnyJson, b: AnyJson): boolean {
  // Same reference or primitive equal
  if (a === b) return true;

  // NaN handling
  if (typeof a === "number" && typeof b === "number") {
    if (Number.isNaN(a) && Number.isNaN(b)) return true;
  }

  return false;
}

/**
 * Generate RFC6902 patch from base -> updated.
 * - Produces deep paths wherever possible.
 * - One op changes one "leaf" (primitive/null), except when type changes fundamentally.
 */
export function generateJsonPatch(base: AnyJson, updated: AnyJson): JsonPatchOp[] {
  const ops: JsonPatchOp[] = [];
  diffIntoOps(base, updated, "", ops);
  return ops;
}

function diffIntoOps(base: AnyJson, updated: AnyJson, path: string, ops: JsonPatchOp[]) {
  // Identical primitives / same reference
  if (valuesEqual(base, updated)) return;

  const baseIsArray = Array.isArray(base);
  const updatedIsArray = Array.isArray(updated);

  // Type changed (object <-> array <-> primitive) => replace whole node
  const baseType = baseIsArray ? "array" : isObject(base) ? "object" : "primitive";
  const updatedType = updatedIsArray ? "array" : isObject(updated) ? "object" : "primitive";
  if (baseType !== updatedType) {
    ops.push({ op: "replace", path: path || "/", value: updated });
    return;
  }

  // Primitive
  if (baseType === "primitive") {
    // base !== updated guaranteed here
    ops.push({ op: "replace", path: path || "/", value: updated });
    return;
  }

  // Array
  if (baseIsArray && updatedIsArray) {
    const minLen = Math.min(base.length, updated.length);

    for (let i = 0; i < minLen; i++) {
      diffIntoOps(base[i], updated[i], joinPath(path, String(i)), ops);
    }

    // Remove extra elements (from end to start to keep indices valid)
    if (base.length > updated.length) {
      for (let i = base.length - 1; i >= updated.length; i--) {
        ops.push({ op: "remove", path: joinPath(path, String(i)) });
      }
    }

    // Add extra elements
    if (updated.length > base.length) {
      for (let i = base.length; i < updated.length; i++) {
        // add at end: "/arr/-" is allowed, but we use explicit index for clarity.
        ops.push({ op: "add", path: joinPath(path, String(i)), value: updated[i] });
      }
    }

    return;
  }

  // Object
  if (isObject(base) && isObject(updated)) {
    const baseKeys = Object.keys(base);
    const updatedKeys = Object.keys(updated);

    const baseSet = new Set(baseKeys);
    const updatedSet = new Set(updatedKeys);

    // removed keys
    for (const k of baseKeys) {
      if (!updatedSet.has(k)) {
        ops.push({ op: "remove", path: joinPath(path, k) });
      }
    }

    // added keys
    for (const k of updatedKeys) {
      if (!baseSet.has(k)) {
        ops.push({ op: "add", path: joinPath(path, k), value: updated[k] });
      }
    }

    // changed keys
    for (const k of updatedKeys) {
      if (baseSet.has(k)) {
        diffIntoOps(base[k], updated[k], joinPath(path, k), ops);
      }
    }

    return;
  }

  // Fallback (shouldn't happen)
  ops.push({ op: "replace", path: path || "/", value: updated });
}
