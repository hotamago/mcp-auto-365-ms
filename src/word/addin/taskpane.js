/* global Office, Word */

(function () {
  "use strict";

  const state = {
    bridgeUrl: window.location.origin,
    sessionId: null,
    token: null,
    docUrl: "",
    docTitle: "",
    polling: false,
    pollController: null,
    heartbeatInterval: null,
  };

  const elements = {
    badge: document.getElementById("connection-badge"),
    docTitle: document.getElementById("doc-title"),
    docUrl: document.getElementById("doc-url"),
    sessionId: document.getElementById("session-id"),
    statusDesc: document.getElementById("bridge-status"),
    logList: document.getElementById("log-list"),
    btnReconnect: document.getElementById("btn-reconnect"),
    btnPing: document.getElementById("btn-ping"),
  };

  function appendLog(text, level = "info") {
    if (!elements.logList) return;
    const now = new Date();
    const timeStr = now.toTimeString().split(" ")[0];
    const row = document.createElement("div");
    row.className = `log-entry log-entry-${level}`;

    const timeSpan = document.createElement("span");
    timeSpan.className = "log-time";
    timeSpan.textContent = timeStr;

    const textSpan = document.createElement("span");
    textSpan.className = "log-text";
    textSpan.textContent = text;

    row.appendChild(timeSpan);
    row.appendChild(textSpan);
    elements.logList.appendChild(row);
    elements.logList.scrollTop = elements.logList.scrollHeight;
  }

  function setStatus(badgeText, badgeClass, desc) {
    if (elements.badge) {
      elements.badge.textContent = badgeText;
      elements.badge.className = `badge badge-${badgeClass}`;
    }
    if (elements.statusDesc && desc) {
      elements.statusDesc.textContent = desc;
    }
  }

  function extractFileNameFromUrl(url) {
    if (!url) return "Tài liệu Word";
    try {
      const parsed = new URL(url);
      const parts = parsed.pathname.split("/").filter(Boolean);
      if (parts.length > 0) {
        const last = decodeURIComponent(parts[parts.length - 1]);
        if (last.endsWith(".docx") || last.endsWith(".doc")) {
          return last;
        }
      }
    } catch {
      // Ignore URL parse error
    }
    return "Tài liệu Word";
  }

  async function resolveDocumentInfo() {
    let url = "";
    try {
      url = Office.context.document.url || "";
    } catch (err) {
      console.warn("Could not read Office.context.document.url directly", err);
    }

    if (!url && Office.context.document.getFilePropertiesAsync) {
      url = await new Promise((resolve) => {
        Office.context.document.getFilePropertiesAsync((result) => {
          if (result && result.status === Office.AsyncResultStatus.Succeeded && result.value) {
            resolve(result.value.url || "");
          } else {
            resolve("");
          }
        });
      });
    }

    state.docUrl = url || window.location.href;
    state.docTitle = extractFileNameFromUrl(state.docUrl);

    if (elements.docTitle) elements.docTitle.textContent = state.docTitle;
    if (elements.docUrl) {
      elements.docUrl.textContent = state.docUrl;
      elements.docUrl.title = state.docUrl;
    }
  }

  async function registerSession() {
    setStatus("Đang kết nối...", "connecting", "Đang đăng ký phiên với local bridge...");
    try {
      const resp = await fetch(`${state.bridgeUrl}/api/sessions/register`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          doc_url: state.docUrl,
          doc_title: state.docTitle,
          platform: Office.context.platform || "web",
          client_id: `word-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
        }),
      });

      if (!resp.ok) {
        throw new Error(`Bridge returned HTTP ${resp.status}`);
      }

      const data = await resp.json();
      state.sessionId = data.session_id;
      state.token = data.token;

      if (elements.sessionId) elements.sessionId.textContent = state.sessionId;
      setStatus("Đã kết nối", "connected", "Sẵn sàng nhận lệnh comment từ harness.");
      appendLog(`Đã kết nối bridge. Session ID: ${state.sessionId}`, "success");

      startHeartbeat();
      startJobPolling();
    } catch (err) {
      setStatus("Lỗi kết nối", "error", `Không kết nối được bridge: ${err.message}`);
      appendLog(`Lỗi kết nối: ${err.message}`, "error");
      setTimeout(registerSession, 5000);
    }
  }

  function startHeartbeat() {
    if (state.heartbeatInterval) clearInterval(state.heartbeatInterval);
    state.heartbeatInterval = setInterval(async () => {
      if (!state.sessionId) return;
      try {
        const resp = await fetch(`${state.bridgeUrl}/api/sessions/${state.sessionId}/heartbeat`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "X-Session-Token": state.token || "",
          },
          body: JSON.stringify({
            doc_url: state.docUrl,
            doc_title: state.docTitle,
          }),
        });
        if (!resp.ok && resp.status === 404) {
          // Session expired on bridge restart; re-register
          appendLog("Phiên bridge đã hết hạn, đang đăng ký lại...", "info");
          registerSession();
        }
      } catch (err) {
        console.warn("Heartbeat error", err);
      }
    }, 5000);
  }

  async function handleAddCommentsJob(job) {
    appendLog(`Nhận lệnh thêm ${job.comments.length} comment từ "${job.author || 'Reviewer'}"...`, "info");
    const results = [];

    await Word.run(async (context) => {
      for (const item of job.comments) {
        const anchor = (item.anchor || "").trim();
        const text = (item.text || "").trim();
        if (!anchor) {
          results.push({ anchor, text, status: "error", error: "Anchor rỗng" });
          continue;
        }

        // Search in document body
        let searchResults = context.document.body.search(anchor, {
          matchCase: false,
          ignorePunct: true,
          ignoreSpace: true,
        });
        searchResults.load(["text"]);
        await context.sync();

        if (searchResults.items.length === 0) {
          // Fallback exact search
          searchResults = context.document.body.search(anchor);
          searchResults.load(["text"]);
          await context.sync();
        }

        if (searchResults.items.length === 0) {
          results.push({
            anchor,
            text,
            status: "error",
            error: `Không tìm thấy cụm từ neo: "${anchor}"`,
          });
          appendLog(`✗ Không tìm thấy: "${anchor.substring(0, 30)}..."`, "error");
        } else {
          const targetRange = searchResults.items[0];
          const comment = targetRange.insertComment(text);
          comment.load(["id", "authorName"]);
          await context.sync();

          results.push({
            anchor,
            text,
            status: "ok",
            comment_id: comment.id || "created",
            author: comment.authorName || job.author,
            matched_text: targetRange.text,
          });
          appendLog(`✓ Đã thêm comment vào: "${targetRange.text.substring(0, 35)}..."`, "success");
        }
      }
    });

    return results;
  }

  async function startJobPolling() {
    if (state.polling) return;
    state.polling = true;

    while (state.sessionId) {
      try {
        state.pollController = new AbortController();
        const resp = await fetch(
          `${state.bridgeUrl}/api/sessions/${state.sessionId}/jobs/poll?timeout=20`,
          {
            headers: { "X-Session-Token": state.token || "" },
            signal: state.pollController.signal,
          }
        );

        if (resp.status === 204) {
          continue;
        }
        if (!resp.ok) {
          if (resp.status === 404) {
            appendLog("Phiên polling 404; đăng ký lại...", "info");
            state.polling = false;
            registerSession();
            return;
          }
          await new Promise((r) => setTimeout(r, 2000));
          continue;
        }

        const job = await resp.json();
        if (!job || !job.job_id) continue;

        let results = [];
        let status = "ok";
        let errorMsg = "";

        try {
          if (job.action === "add_comments") {
            results = await handleAddCommentsJob(job);
          } else {
            throw new Error(`Hành động không được hỗ trợ: ${job.action}`);
          }
        } catch (jobErr) {
          status = "error";
          errorMsg = jobErr.message || String(jobErr);
          appendLog(`Lỗi xử lý job: ${errorMsg}`, "error");
        }

        // Post result back
        await fetch(`${state.bridgeUrl}/api/sessions/${state.sessionId}/jobs/${job.job_id}/result`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "X-Session-Token": state.token || "",
          },
          body: JSON.stringify({
            status,
            results,
            error: errorMsg,
          }),
        });
      } catch (pollErr) {
        if (pollErr.name === "AbortError") break;
        await new Promise((r) => setTimeout(r, 3000));
      }
    }
    state.polling = false;
  }

  function setupButtons() {
    if (elements.btnReconnect) {
      elements.btnReconnect.addEventListener("click", () => {
        appendLog("Đang làm mới kết nối theo yêu cầu...", "info");
        registerSession();
      });
    }
    if (elements.btnPing) {
      elements.btnPing.addEventListener("click", async () => {
        try {
          const resp = await fetch(`${state.bridgeUrl}/api/status`);
          const data = await resp.json();
          appendLog(`Bridge OK (Uptime: ${Math.round(data.uptime)}s, Sessions: ${data.active_sessions})`, "success");
        } catch (err) {
          appendLog(`Bridge không phản hồi: ${err.message}`, "error");
        }
      });
    }
  }

  Office.onReady(async (info) => {
    appendLog(`Office.js đã sẵn sàng (Host: ${info.host}, Platform: ${info.platform})`, "info");
    setupButtons();

    if (info.host === Office.HostType.Word) {
      await resolveDocumentInfo();
      registerSession();
    } else {
      setStatus("Không hỗ trợ", "error", `Host ${info.host} không phải là Microsoft Word.`);
      appendLog(`Host ${info.host} không phải là Word.`, "error");
    }
  });
})();
