// ==UserScript==
// @name         M365 Copilot Chat Exporter
// @namespace    https://github.com/ai-experiments
// @version      4.1
// @description  Export Microsoft 365 Copilot conversations as ChatGPT-compatible conversations.json
// @license      MIT
// @author       ingo
// @match        https://m365.cloud.microsoft/
// @match        https://m365.cloud.microsoft/chat*
// @match        https://microsoft365.com/chat*
// @match        https://www.microsoft365.com/chat*
// @icon         data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Cdefs%3E%3ClinearGradient id='g' x1='0' y1='0' x2='1' y2='1'%3E%3Cstop offset='0%25' stop-color='%23a855f7'/%3E%3Cstop offset='100%25' stop-color='%236366f1'/%3E%3C/linearGradient%3E%3C/defs%3E%3Crect width='32' height='32' rx='6' fill='url(%23g)'/%3E%3Cpath d='M9 11h14M9 16h10M9 21h12' stroke='white' stroke-width='2' stroke-linecap='round'/%3E%3Cpath d='M22 18l3 3-3 3' stroke='%2322d3ee' stroke-width='2' stroke-linecap='round' stroke-linejoin='round' fill='none'/%3E%3C/svg%3E
// @grant        none
// @run-at       document-end
// ==/UserScript==

(function () {
  "use strict";

  // ── State ──────────────────────────────────────────────────────────

  const conversations = new Map();
  const rawCaptures = [];
  let isFetchingAll = false;

  const SKIP_MESSAGE_TYPES = new Set([
    "CrossPluginGroundingData",
    "Internal",
    "InternalSuggestions",
    "InternalLoaderMessage",
    "InternalSearchResult",
    "InternalSearchQuery",
    "Suggestion",
    "RenderCardRequest",
    "GenerateContentQuery",
    "AdsQuery",
  ]);

  // ── Date range helpers ──────────────────────────────────────────

  /**
   * Returns { from: Date, to: Date } for the selected date range preset,
   * or null if "all" is selected.  Timestamps are midnight-based in local tz.
   */
  function getDateRange() {
    const sel = document.getElementById("copilot-date-range");
    if (!sel) return null;
    const preset = sel.value;
    if (preset === "all") return null;

    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);

    if (preset === "custom") {
      const fromEl = document.getElementById("copilot-date-from");
      const toEl = document.getElementById("copilot-date-to");
      const from = fromEl?.value ? new Date(fromEl.value) : null;
      const to = toEl?.value ? new Date(toEl.value) : null;
      if (!from && !to) return null;
      // "to" is inclusive: advance to the next day
      const toEnd = to ? new Date(to.getTime() + 86400000) : tomorrow;
      return { from: from || new Date(0), to: toEnd };
    }

    let from;
    switch (preset) {
      case "today":
        from = today;
        break;
      case "week": {
        from = new Date(today);
        from.setDate(from.getDate() - 7);
        break;
      }
      case "month": {
        from = new Date(today);
        from.setMonth(from.getMonth() - 1);
        break;
      }
      case "year": {
        from = new Date(today);
        from.setFullYear(from.getFullYear() - 1);
        break;
      }
      default:
        return null;
    }
    return { from, to: tomorrow };
  }

  /**
   * Check whether a conversation falls within the active date range.
   * createTimeUtc from the Substrate API is in Unix milliseconds.
   */
  function isInDateRange(conv, range) {
    if (!range) return true;
    const ts = conv.createTimeUtc;
    if (!ts) return true; // keep conversations with unknown dates
    const d = new Date(typeof ts === "number" && ts > 1e12 ? ts : ts);
    return d >= range.from && d < range.to;
  }

  const SUBSTRATE_BASE = "https://substrate.office.com/m365Copilot";
  const DEFAULT_VARIANTS =
    "feature.EnableLastMessageForGetChats,feature.EnableMRUAgents,feature.EnableHasLoopPages,feature.EnableIsInputControlInGptItem";

  // ── MSAL token extraction (adapted from ganyuke/copilot-exporter) ─

  const getCookie = (key) =>
    document.cookie.match(`(^|;)\\s*${key}\\s*=\\s*([^;]+)`)?.pop() || "";

  function base64DecToArr(base64String) {
    let s = base64String.replace(/-/g, "+").replace(/_/g, "/");
    switch (s.length % 4) {
      case 2: s += "=="; break;
      case 3: s += "="; break;
    }
    const bin = atob(s);
    return Uint8Array.from(bin, (c) => c.codePointAt(0) || 0);
  }

  function toArrayBuffer(bufferLike) {
    return Uint8Array.from(bufferLike).buffer;
  }

  async function deriveKey(baseKey, nonce, context) {
    return crypto.subtle.deriveKey(
      { name: "HKDF", salt: toArrayBuffer(nonce), hash: "SHA-256", info: new TextEncoder().encode(context) },
      baseKey,
      { name: "AES-GCM", length: 256 },
      false,
      ["encrypt", "decrypt"]
    );
  }

  async function decryptPayload(baseKey, nonce, context, encryptedData) {
    const encoded = base64DecToArr(encryptedData);
    const derived = await deriveKey(baseKey, base64DecToArr(nonce), context);
    const decrypted = await crypto.subtle.decrypt(
      { name: "AES-GCM", iv: new Uint8Array(12) },
      derived,
      toArrayBuffer(encoded)
    );
    return new TextDecoder().decode(decrypted);
  }

  async function getEncryptionCookie() {
    const raw = decodeURIComponent(getCookie("msal.cache.encryption"));
    let parsed;
    try { parsed = JSON.parse(raw); } catch { throw new Error("Failed to parse msal.cache.encryption cookie"); }
    if (!parsed?.key || !parsed?.id) throw new Error("No encryption cookie found");
    return {
      id: parsed.id,
      key: await crypto.subtle.importKey("raw", toArrayBuffer(base64DecToArr(parsed.key)), "HKDF", false, ["deriveKey"]),
    };
  }

  function getMsalIds() {
    const el = document.getElementById("identity");
    if (!el?.textContent) throw new Error("Missing #identity element in page");
    const { objectId, tenantId } = JSON.parse(el.textContent);
    return {
      localAccountId: objectId,
      tenantId,
      homeAccountId: `${objectId}.${tenantId}`,
      clientId: "c0ab8ce9-e9a0-42e7-b064-33d422df41f1",
    };
  }

  async function getAccessToken(msalIds) {
    const cookie = await getEncryptionCookie();
    const { homeAccountId, tenantId, clientId } = msalIds;
    const scopes = ["https://substrate.office.com/sydney/.default"];
    const lsKey = `${homeAccountId}-login.windows.net-accesstoken-${clientId}-${tenantId}-${scopes.join(" ")}--`;
    const stored = localStorage.getItem(lsKey);
    if (!stored) throw new Error("Missing MSAL access token in localStorage");
    const payload = JSON.parse(stored);
    const decrypted = await decryptPayload(cookie.key, payload.nonce, clientId, payload.data);
    return JSON.parse(decrypted).secret;
  }

  async function getTokenAndIds() {
    const msalIds = getMsalIds();
    const token = await getAccessToken(msalIds);
    return { token, ...msalIds };
  }

  // ── Substrate API handlers ─────────────────────────────────────────

  function handleGetConversation(data) {
    const convId = data.conversationId;
    if (!convId) return;

    const visibleMessages = (data.messages || []).filter((m) => {
      if (SKIP_MESSAGE_TYPES.has(m.messageType)) return false;
      if (m.author === "system") return false;
      if (!m.text && !m.adaptiveCards?.length) return false;
      return true;
    });

    if (visibleMessages.length === 0) return;

    conversations.set(convId, {
      conversationId: convId,
      chatName: data.chatName || "",
      createTimeUtc: data.createTimeUtc,
      updateTimeUtc: data.updateTimeUtc,
      tone: data.tone || "",
      isLegacyWebChat: data.isLegacyWebChat || false,
      messages: visibleMessages,
    });

    updateBadge();
  }

  function handleGetChats(data) {
    const chats = data.chats || [];
    for (const chat of chats) {
      const convId = chat.conversationId;
      if (!convId) continue;
      if (!conversations.has(convId)) {
        conversations.set(convId, {
          conversationId: convId,
          chatName: chat.chatName || "",
          createTimeUtc: chat.createTimeUtc,
          updateTimeUtc: chat.updateTimeUtc,
          tone: chat.tone || "",
          isLegacyWebChat: chat.isLegacyWebChat || false,
          messages: [],
        });
      }
    }
    console.log(
      `[Copilot Export] Chat list: ${chats.length} conversations (${conversations.size} total known)`
    );
    updateBadge();
    return data;
  }

  // ── Passive response interceptors (for badge + raw export) ────────

  const originalFetch = window.fetch;

  window.fetch = async function (...args) {
    const url = typeof args[0] === "string" ? args[0] : args[0]?.url || "";
    const response = await originalFetch.apply(this, args);

    if (!url.includes("substrate.office.com") && !url.includes("m365.cloud.microsoft")) {
      return response;
    }

    const clone = response.clone();
    clone.text().then((text) => {
      if (!text || text.length < 20) return;
      try {
        const json = JSON.parse(text);
        rawCaptures.push({
          url: url.substring(0, 500),
          status: response.status,
          timestamp: new Date().toISOString(),
          byteLength: text.length,
          data: json,
        });
        if (url.includes("GetConversation")) handleGetConversation(json);
        else if (url.includes("GetChats")) handleGetChats(json);
      } catch { /* not JSON */ }
    }).catch(() => {});

    return response;
  };

  const origXHROpen = XMLHttpRequest.prototype.open;
  const origXHRSend = XMLHttpRequest.prototype.send;

  XMLHttpRequest.prototype.open = function (method, url, ...rest) {
    this._captureUrl = url;
    return origXHROpen.call(this, method, url, ...rest);
  };

  XMLHttpRequest.prototype.send = function (...args) {
    this.addEventListener("load", function () {
      const url = this._captureUrl || "";
      if (!url.includes("substrate.office.com") && !url.includes("m365.cloud.microsoft")) return;
      try {
        const json = JSON.parse(this.responseText);
        rawCaptures.push({
          url: url.substring(0, 500),
          status: this.status,
          timestamp: new Date().toISOString(),
          byteLength: this.responseText.length,
          data: json,
        });
        if (url.includes("GetConversation")) handleGetConversation(json);
        else if (url.includes("GetChats")) handleGetChats(json);
      } catch { /* not JSON */ }
    });
    return origXHRSend.apply(this, args);
  };

  // ── Fetch All automation ──────────────────────────────────────────

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function setStatus(text) {
    const el = document.getElementById("copilot-export-status");
    if (el) el.textContent = text;
  }

  /**
   * Call a Substrate endpoint with proper auth headers.
   */
  async function substrateGet(auth, endpoint, params, includeVariants) {
    const requestJson = JSON.stringify(params);
    const variantsSuffix = includeVariants
      ? `&variants=${encodeURIComponent(DEFAULT_VARIANTS)}`
      : "";
    const url = `${SUBSTRATE_BASE}/${endpoint}?request=${encodeURIComponent(requestJson)}${variantsSuffix}`;

    const headers = {
      authorization: `Bearer ${auth.token}`,
      "content-type": "application/json",
      "x-anchormailbox": `Oid:${auth.localAccountId}@${auth.tenantId}`,
      "x-clientrequestid": crypto.randomUUID().replace(/-/g, ""),
      "x-routingparameter-sessionkey": auth.localAccountId,
      "x-scenario": "OfficeWebIncludedCopilot",
    };

    const resp = await fetch(url, { method: "GET", headers });

    if (!resp.ok) {
      throw new Error(`${endpoint} returned ${resp.status}`);
    }
    return resp.json();
  }

  /**
   * Fetch all chat IDs by paginating through GetChats.
   */
  async function fetchAllChatIds(auth) {
    const allChats = [];
    let syncState = null;
    let page = 0;

    while (true) {
      page++;
      setStatus(`Fetching chat list page ${page}...`);

      const params = {
        source: "officeweb",
        traceId: crypto.randomUUID(),
        threadType: "bizchat",
        MaxReturnedChatsCount: 50,
        mergeWorkWebChats: true,
        includeChatsWithHarmfulContentProtectionDisabled: true,
      };
      if (syncState) {
        params.syncState = syncState;
      }

      const data = await substrateGet(auth, "GetChats", params, true);
      const chats = data.chats || [];
      allChats.push(...chats);
      handleGetChats(data);

      console.log(
        `[Copilot Export] GetChats page ${page}: ${chats.length} chats (${allChats.length} total)`
      );

      syncState = data.syncState || null;
      if (chats.length === 0 || !syncState) break;

      await sleep(500);
    }

    return allChats;
  }

  /**
   * Fetch full conversation content for a single chat.
   */
  async function fetchConversation(auth, conversationId) {
    const data = await substrateGet(auth, "GetConversation", {
      conversationId,
      source: "officeweb",
      traceId: crypto.randomUUID().replace(/-/g, ""),
    }, false);
    handleGetConversation(data);
    return data;
  }

  /**
   * Main "Fetch All" workflow.
   */
  async function doFetchAll() {
    if (isFetchingAll) return;
    isFetchingAll = true;
    const btn = document.getElementById("copilot-btn-fetchall");
    if (btn) { btn.disabled = true; btn.textContent = "Fetching..."; }

    try {
      setStatus("Acquiring auth token...");
      const auth = await getTokenAndIds();
      console.log(`[Copilot Export] Auth acquired for ${auth.localAccountId}`);

      // Step 1: Get all chat IDs
      const allChats = await fetchAllChatIds(auth);
      console.log(`[Copilot Export] Found ${allChats.length} total conversations`);

      // Step 1b: Apply date range filter
      const range = getDateRange();
      const filteredChats = range
        ? allChats.filter((c) => isInDateRange(c, range))
        : allChats;

      if (range) {
        console.log(`[Copilot Export] Date filter: ${filteredChats.length}/${allChats.length} conversations in range`);
        setStatus(`${filteredChats.length} of ${allChats.length} conversations match date filter`);
        await sleep(800);
      }

      // Step 2: Fetch each conversation that we don't already have messages for
      const toFetch = filteredChats.filter((c) => {
        const existing = conversations.get(c.conversationId);
        return !existing || existing.messages.length === 0;
      });

      console.log(`[Copilot Export] Need to fetch ${toFetch.length} conversations (${filteredChats.length - toFetch.length} already loaded)`);

      let fetched = 0;
      let errors = 0;

      for (const chat of toFetch) {
        fetched++;
        setStatus(`Fetching ${fetched}/${toFetch.length}: ${(chat.chatName || "").substring(0, 40)}...`);

        try {
          await fetchConversation(auth, chat.conversationId);
        } catch (e) {
          errors++;
          console.warn(`[Copilot Export] Failed to fetch ${chat.conversationId}: ${e.message}`);
        }

        await sleep(500);
        updateBadge();
      }

      const withMessages = Array.from(conversations.values()).filter(
        (c) => c.messages && c.messages.length > 0
      );
      setStatus(
        `Done! ${withMessages.length} conversations loaded` +
          (errors > 0 ? ` (${errors} errors)` : "")
      );
    } catch (e) {
      setStatus(`Error: ${e.message}`);
      console.error("[Copilot Export] Fetch all failed:", e);
    } finally {
      isFetchingAll = false;
      if (btn) { btn.disabled = false; btn.textContent = "Fetch All Conversations"; }
    }
  }

  // ── ChatGPT format converter ──────────────────────────────────────

  function toUnixSeconds(ts) {
    if (!ts) return null;
    if (typeof ts === "number") return ts > 1e12 ? ts / 1000 : ts;
    try { return new Date(ts).getTime() / 1000; } catch { return null; }
  }

  function buildConversationsJson(range) {
    const output = [];

    for (const [convId, conv] of conversations) {
      if (!conv.messages || conv.messages.length === 0) continue;
      if (!isInDateRange(conv, range)) continue;

      const firstTs =
        toUnixSeconds(conv.createTimeUtc) ||
        toUnixSeconds(conv.messages[0]?.createdAt);
      const lastTs =
        toUnixSeconds(conv.updateTimeUtc) ||
        toUnixSeconds(conv.messages[conv.messages.length - 1]?.createdAt);

      const mapping = {};

      const rootId = "client-created-root";
      mapping[rootId] = { id: rootId, message: null, parent: null, children: [] };

      const systemId = crypto.randomUUID();
      mapping[systemId] = {
        id: systemId,
        message: {
          id: systemId,
          author: { role: "system", name: null, metadata: {} },
          create_time: firstTs,
          update_time: null,
          content: { content_type: "text", parts: [""] },
          status: "finished_successfully",
          end_turn: true,
          weight: 1.0,
          metadata: {},
          recipient: "all",
          channel: null,
        },
        parent: rootId,
        children: [],
      };
      mapping[rootId].children.push(systemId);

      let prevId = systemId;
      let title = conv.chatName || null;

      for (const msg of conv.messages) {
        const created = toUnixSeconds(msg.createdAt);
        const text = msg.text || "";

        let role;
        if (msg.author === "user") {
          role = "user";
          if (!title && text) title = text.substring(0, 100).trim();
        } else if (msg.author === "bot") {
          role = "assistant";
        } else {
          continue;
        }

        const nodeId = msg.messageId || crypto.randomUUID();

        mapping[nodeId] = {
          id: nodeId,
          message: {
            id: nodeId,
            author: { role, name: null, metadata: {} },
            create_time: created,
            update_time: null,
            content: { content_type: "text", parts: [text] },
            status: "finished_successfully",
            end_turn: role === "assistant",
            weight: 1.0,
            metadata: {
              copilot_app_class: msg.contentOrigin || "",
              copilot_session_id: convId,
              copilot_message_id: msg.messageId || "",
              copilot_request_id: msg.requestId || "",
            },
            recipient: "all",
            channel: null,
          },
          parent: prevId,
          children: [],
        };

        mapping[prevId].children.push(nodeId);
        prevId = nodeId;
      }

      output.push({
        title: title || "Copilot Chat",
        create_time: firstTs,
        update_time: lastTs,
        mapping,
        moderation_results: [],
        current_node: prevId,
        plugin_ids: null,
        conversation_id: convId,
        conversation_template_id: null,
        gizmo_id: null,
        gizmo_type: null,
        is_archived: false,
        is_starred: null,
        safe_urls: [],
        blocked_urls: [],
        default_model_slug: "copilot",
        conversation_origin: null,
        is_read_only: null,
        voice: null,
        async_status: null,
        disabled_tool_ids: [],
        is_do_not_remember: false,
        memory_scope: null,
        context_scopes: null,
        sugar_item_id: null,
        sugar_item_visible: false,
        pinned_time: null,
        is_study_mode: false,
        owner: null,
        id: convId,
      });
    }

    output.sort((a, b) => (b.create_time || 0) - (a.create_time || 0));
    return output;
  }

  // ── Download helpers ───────────────────────────────────────────────

  function downloadJson(data, filename) {
    // Replace Unicode line/paragraph separators (U+2028, U+2029) with \n
    const jsonStr = JSON.stringify(data, null, 2).replace(/[\u2028\u2029]/g, "\n");
    const blob = new Blob([jsonStr], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function doExportConverted() {
    const range = getDateRange();
    const withMessages = Array.from(conversations.values()).filter(
      (c) => c.messages && c.messages.length > 0 && isInDateRange(c, range)
    );

    if (withMessages.length === 0) {
      alert(
        range
          ? "No conversations with content in the selected date range.\n\nTry a wider date range or use Fetch All first."
          : "No conversation content captured yet.\n\n" +
              'Use "Fetch All Conversations" to load everything first.'
      );
      return;
    }

    const result = buildConversationsJson(range);
    const totalMsgs = withMessages.reduce((s, c) => s + c.messages.length, 0);

    downloadJson(
      { conversations: result },
      `copilot_conversations_${new Date().toISOString().slice(0, 10)}.json`
    );

    console.log(
      `[Copilot Export] Exported ${result.length} conversations with ${totalMsgs} messages`
    );
  }

  function doExportRaw() {
    if (rawCaptures.length === 0) {
      alert("No API responses captured yet.");
      return;
    }
    downloadJson(
      rawCaptures,
      `copilot_raw_capture_${new Date().toISOString().slice(0, 10)}.json`
    );
  }

  // ── Floating UI ───────────────────────────────────────────────────

  function updateBadge() {
    const badge = document.getElementById("copilot-export-badge");
    if (badge) {
      const withMessages = Array.from(conversations.values()).filter(
        (c) => c.messages && c.messages.length > 0
      );
      const totalMsgs = withMessages.reduce((s, c) => s + c.messages.length, 0);
      const pending = conversations.size - withMessages.length;

      let text = `${withMessages.length} chats / ${totalMsgs} msgs captured`;
      if (pending > 0) text += ` (${pending} not yet loaded)`;
      badge.textContent = text;
    }
  }

  function createUI() {
    const container = document.createElement("div");
    container.id = "copilot-export-ui";
    container.innerHTML = `
      <style>
        #copilot-export-ui {
          position: fixed;
          bottom: 16px;
          right: 16px;
          z-index: 999999;
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
          font-size: 13px;
        }
        #copilot-export-title-row {
          display: flex;
          align-items: center;
          justify-content: space-between;
          margin-bottom: 8px;
        }
        #copilot-export-minimize-btn {
          background: transparent;
          border: none;
          color: #555;
          cursor: pointer;
          font-size: 16px;
          padding: 0 0 0 8px;
          line-height: 1;
          opacity: 0.5;
          transition: opacity 0.15s;
        }
        #copilot-export-minimize-btn:hover {
          opacity: 1;
          color: #9d8aff;
        }
        #copilot-export-icon {
          display: none;
          width: 48px;
          height: 48px;
          background: linear-gradient(135deg, #a855f7 0%, #6366f1 100%);
          border-radius: 50%;
          cursor: pointer;
          box-shadow: 0 4px 20px rgba(0,0,0,0.4);
          align-items: center;
          justify-content: center;
          font-size: 24px;
          color: white;
          transition: transform 0.2s;
        }
        #copilot-export-icon:hover {
          transform: scale(1.1);
        }
        #copilot-export-ui.minimized #copilot-export-panel {
          display: none;
        }
        #copilot-export-ui.minimized #copilot-export-icon {
          display: flex;
        }
        #copilot-export-panel {
          background: #1a1a2e;
          color: #e0e0e0;
          border-radius: 10px;
          padding: 12px 16px;
          box-shadow: 0 4px 20px rgba(0,0,0,0.4);
          min-width: 260px;
        }
        #copilot-export-panel .title {
          font-weight: 600;
          font-size: 13px;
          color: #9d8aff;
        }
        #copilot-export-badge {
          font-size: 12px;
          color: #8be9fd;
          margin-bottom: 6px;
          font-variant-numeric: tabular-nums;
          line-height: 1.4;
        }
        #copilot-export-status {
          font-size: 11px;
          color: #50fa7b;
          margin-bottom: 10px;
          line-height: 1.4;
          min-height: 15px;
        }
        #copilot-export-panel button {
          display: block;
          width: 100%;
          padding: 7px 10px;
          margin-bottom: 6px;
          border: 1px solid #333;
          border-radius: 6px;
          background: #16213e;
          color: #e0e0e0;
          cursor: pointer;
          font-size: 12px;
          text-align: left;
        }
        #copilot-export-panel button:hover:not(:disabled) {
          background: #1f3460;
          border-color: #9d8aff;
        }
        #copilot-export-panel button:disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }
        #copilot-export-panel button.primary {
          background: #1f3460;
          border-color: #9d8aff;
          font-weight: 600;
        }
        #copilot-export-panel button.primary:hover:not(:disabled) {
          background: #2a4a80;
        }
        #copilot-export-panel .hint {
          font-size: 11px;
          color: #666;
          margin-top: 8px;
          line-height: 1.4;
        }
        #copilot-date-row {
          margin-bottom: 8px;
        }
        #copilot-date-row label {
          font-size: 11px;
          color: #aaa;
          display: block;
          margin-bottom: 3px;
        }
        #copilot-date-row select,
        #copilot-date-row input[type="date"] {
          background: #16213e;
          color: #e0e0e0;
          border: 1px solid #333;
          border-radius: 4px;
          padding: 4px 6px;
          font-size: 12px;
          font-family: inherit;
        }
        #copilot-date-row select {
          width: 100%;
        }
        #copilot-custom-dates {
          display: none;
          margin-top: 4px;
          gap: 6px;
        }
        #copilot-custom-dates.visible {
          display: flex;
        }
        #copilot-custom-dates input[type="date"] {
          flex: 1;
          min-width: 0;
        }
      </style>
      <div id="copilot-export-icon" title="Open Copilot Exporter">📥</div>
      <div id="copilot-export-panel">
        <div id="copilot-export-title-row">
          <div class="title">Copilot Chat Exporter v4.1</div>
          <button id="copilot-export-minimize-btn" title="Minimize" style="width:30px">−</button>
        </div>
        <div id="copilot-export-badge">Waiting for data...</div>
        <div id="copilot-export-status"></div>
        <div id="copilot-date-row">
          <label>Date range</label>
          <select id="copilot-date-range">
            <option value="all">All time</option>
            <option value="today">Today</option>
            <option value="week">Last 7 days</option>
            <option value="month">Last 30 days</option>
            <option value="year">Last year</option>
            <option value="custom">Custom...</option>
          </select>
          <div id="copilot-custom-dates">
            <input type="date" id="copilot-date-from" title="From date">
            <input type="date" id="copilot-date-to" title="To date">
          </div>
        </div>
        <button id="copilot-btn-fetchall" class="primary">Fetch All Conversations</button>
        <button id="copilot-btn-export">Export conversations.json</button>
        <button id="copilot-btn-raw">Export raw API captures</button>
        <div class="hint">Click Fetch All to load all conversations<br>directly from the API. <a href="https://github.com/ingo/m365_copilot_chat_exporter" target="_blank" style="color: #9d8aff; text-decoration: none;">About</a></div>
      </div>
    `;

    document.body.appendChild(container);
    document.getElementById("copilot-btn-fetchall").addEventListener("click", doFetchAll);
    document.getElementById("copilot-btn-export").addEventListener("click", doExportConverted);
    document.getElementById("copilot-btn-raw").addEventListener("click", doExportRaw);

    document.getElementById("copilot-date-range").addEventListener("change", (e) => {
      const customRow = document.getElementById("copilot-custom-dates");
      customRow.classList.toggle("visible", e.target.value === "custom");
    });

    // Minimize/maximize toggle
    const ui = document.getElementById("copilot-export-ui");
    const minimizeBtn = document.getElementById("copilot-export-minimize-btn");
    const icon = document.getElementById("copilot-export-icon");

    minimizeBtn.addEventListener("click", () => {
      ui.classList.add("minimized");
    });

    icon.addEventListener("click", () => {
      ui.classList.remove("minimized");
    });
  }

  // ── Init ──────────────────────────────────────────────────────────
  console.log("[Copilot Export v4.1] Loaded.");
  createUI();
})();
