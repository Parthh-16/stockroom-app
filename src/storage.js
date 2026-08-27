// Polyfills the window.storage API (originally a Claude.ai-artifact-only
// feature) using Supabase, so multiple people/devices share the same
// inventory, invoices, returns, and login data in real time.
//
// App.jsx itself never needs to know about this change — it only ever
// talks to window.storage, exactly as before.

import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://rkbvcqlsdfqpkwfygnhs.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_sm796VS24BsjC1WtoX59sA_KUoys67k";

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
const TABLE = "stockroom_kv";

async function get(key) {
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
