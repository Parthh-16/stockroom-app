// Polyfills the window.storage API (originally a Claude.ai-artifact-only
// feature).
//
// Most data (inventory, invoices, returns, login credentials) is stored
// in Supabase, so multiple people/devices share the same data in real
// time. The one exception is "am I currently logged in on THIS device"
// (SESSION_KEY below) — that stays in localStorage, per device, exactly
// like a normal app. Otherwise logging in on one phone would silently
// log in every other device too, and any brief network hiccup while
// reading a shared value could bounce someone back to the login screen.
//
// App.jsx itself never needs to know about this split — it only ever
// talks to window.storage, exactly as before.

import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://rkbvcqlsdfqpkwfygnhs.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_sm796VS24BsjC1WtoX59sA_KUoys67k";

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
const TABLE = "stockroom_kv";

// Keys that should stay local to this device/browser instead of shared.
// Uses sessionStorage (not localStorage) on purpose: this keeps you
// logged in through a page refresh, but requires logging in again once
// the browser tab/window is fully closed — the expected behavior for a
// shared work app, instead of staying logged in forever on any device.
const LOCAL_ONLY_KEYS = new Set(["stockroom:session"]);
const LOCAL_PREFIX = "stockroom_app_local::";

function localGet(key) {
  try {
    const raw = window.sessionStorage.getItem(LOCAL_PREFIX + key);
    if (raw === null) return null;
    return { key, value: raw, shared: false };
  } catch (e) {
    return null;
  }
}

function localSet(key, value) {
  try {
    window.sessionStorage.setItem(LOCAL_PREFIX + key, value);
    return { key, value, shared: false };
  } catch (e) {
    return null;
  }
}

function localDelete(key) {
  try {
    const existed = window.sessionStorage.getItem(LOCAL_PREFIX + key) !== null;
    window.sessionStorage.removeItem(LOCAL_PREFIX + key);
    return { key, deleted: existed, shared: false };
  } catch (e) {
    return null;
  }
}

async function get(key) {
  if (LOCAL_ONLY_KEYS.has(key)) return localGet(key);
  try {
    const { data, error } = await supabase
      .from(TABLE)
      .select("value")
      .eq("key", key)
      .maybeSingle();
    if (error) throw error;
    if (!data) return null;
    return { key, value: data.value, shared: true };
  } catch (e) {
    console.error("storage.get failed:", e);
    return null;
  }
}

async function set(key, value) {
  if (LOCAL_ONLY_KEYS.has(key)) return localSet(key, value);
  try {
    const { error } = await supabase
      .from(TABLE)
      .upsert({ key, value, updated_at: new Date().toISOString() }, { onConflict: "key" });
    if (error) throw error;
    return { key, value, shared: true };
  } catch (e) {
    console.error("storage.set failed:", e);
    return null;
  }
}

async function del(key) {
  if (LOCAL_ONLY_KEYS.has(key)) return localDelete(key);
  try {
    const { data: existing } = await supabase.from(TABLE).select("key").eq("key", key).maybeSingle();
    const existed = !!existing;
    const { error } = await supabase.from(TABLE).delete().eq("key", key);
    if (error) throw error;
    return { key, deleted: existed, shared: true };
  } catch (e) {
    console.error("storage.delete failed:", e);
    return null;
  }
}

async function list(prefix = "") {
  try {
    const { data, error } = await supabase
      .from(TABLE)
      .select("key")
      .ilike("key", `${prefix}%`);
    if (error) throw error;
    return { keys: (data || []).map((row) => row.key), prefix, shared: true };
  } catch (e) {
    console.error("storage.list failed:", e);
    return { keys: [], prefix, shared: true };
  }
}

if (typeof window !== "undefined" && !window.storage) {
  window.storage = { get, set, delete: del, list };
}
