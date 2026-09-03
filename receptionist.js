/**
 * Smile Well IV reception desk.
 * Staff-only page — do not link from the public membership form.
 */
const WEBHOOK_URL =
  "https://script.google.com/macros/s/AKfycbxh_7LqirydQHGSSxF4r3jcs5WKEDMHAT-3BiQIg5wmwCpsS9PjrCijozryZYpHOMtzeg/exec";

const firebaseConfig = {
  apiKey: "AIzaSyDgV0ZN5h1MWzQWNwNqe-ZJmy2aBWL8diI",
  authDomain: "smile-well-34579.firebaseapp.com",
  projectId: "smile-well-34579",
  storageBucket: "smile-well-34579.firebasestorage.app",
  messagingSenderId: "462381026785",
  appId: "1:462381026785:web:7e35027eb715c5199616a0",
  measurementId: "G-B67NQ5PX1M",
};

if (typeof firebase === "undefined" || !firebase.initializeApp || !firebase.auth) {
  const loadingEl = document.getElementById("auth-loading");
  if (loadingEl) {
    loadingEl.innerHTML =
      '<p class="text-sm text-red-700">Could not load sign-in. Refresh the page.</p>';
  }
  throw new Error("Firebase Auth SDK missing");
}

firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();

const authLoading = document.getElementById("auth-loading");
const deskApp = document.getElementById("desk-app");
const signedInEmail = document.getElementById("signed-in-email");
const signOutBtn = document.getElementById("sign-out-btn");
const searchInput = document.getElementById("client-search");
const listEl = document.getElementById("client-list");
const listMeta = document.getElementById("list-meta");
const listPane = document.getElementById("list-pane");
const detailPane = document.getElementById("detail-pane");
const detailBody = document.getElementById("detail-body");
const refreshBtn = document.getElementById("refresh-btn");

let clients = [];
let selectedId = "";
let loading = false;
let saving = false;
let lastError = "";
let signedIn = false;

function digitsOnly(value) {
  return String(value || "").replace(/\D/g, "");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function parseAmount(raw) {
  const text = String(raw || "").trim();
  if (!text) return "";
  const cleaned = text.replace(/CAD/gi, "").replace(/[$,\s]/g, "");
  if (!/^\d+(\.\d{1,2})?$/.test(cleaned)) return null;
  const n = Number(cleaned);
  if (!Number.isFinite(n) || n < 0 || n > 9999999.99) return null;
  return n.toFixed(2);
}

function formatCAD(amount) {
  if (amount === "" || amount == null) return "";
  const n = Number(amount);
  if (!Number.isFinite(n)) return String(amount);
  return new Intl.NumberFormat("en-CA", {
    style: "currency",
    currency: "CAD",
  }).format(n);
}

function statusClass(status) {
  const key = String(status || "").toLowerCase();
  if (key === "active") return "bg-emerald-50 text-emerald-800 border-emerald-200";
  if (key === "contacted") return "bg-amber-50 text-amber-900 border-amber-200";
  if (key === "expired") return "bg-stone-100 text-stone-600 border-stone-200";
  if (key === "closed") return "bg-rose-50 text-rose-800 border-rose-200";
  return "bg-navy-soft text-navy border-navy/15";
}

function parseInternalNotes(notes) {
  const result = { dob: "", signature: "", heardFrom: "", extras: [] };
  String(notes || "")
    .split(" | ")
    .map((part) => part.trim())
    .filter(Boolean)
    .forEach((part) => {
      if (part.startsWith("DOB: ")) result.dob = part.slice(5);
      else if (part.startsWith("Signed: ")) result.signature = part.slice(8);
      else if (part.startsWith("Heard from: ")) result.heardFrom = part.slice(12);
      else result.extras.push(part);
    });
  return result;
}

function matchesQuery(client, query) {
  const q = query.trim().toLowerCase();
  if (!q) return true;
  const name = String(client.fullName || "").toLowerCase();
  const id = String(client.id || "").toLowerCase();
  const email = String(client.email || "").toLowerCase();
  const phone = digitsOnly(client.phone);
  const qDigits = digitsOnly(q);
  if (name.includes(q) || id.includes(q) || email.includes(q)) return true;
  return qDigits.length >= 3 && phone.includes(qDigits);
}

function filteredClients() {
  return clients.filter((client) => matchesQuery(client, searchInput.value));
}

function selectedClient() {
  return clients.find((client) => client.id === selectedId) || null;
}

function isWide() {
  return window.matchMedia("(min-width: 1024px)").matches;
}

function showList() {
  if (isWide()) {
    listPane.classList.remove("hidden");
    detailPane.classList.remove("hidden");
    return;
  }
  if (selectedId) {
    listPane.classList.add("hidden");
    detailPane.classList.remove("hidden");
  } else {
    listPane.classList.remove("hidden");
    detailPane.classList.add("hidden");
  }
}

function renderList() {
  const rows = filteredClients();
  const total = clients.length;
  if (lastError) {
    listMeta.textContent = "";
  } else if (loading) {
    listMeta.textContent = "Loading members…";
  } else if (!total) {
    listMeta.textContent = "No members yet";
  } else if (!rows.length) {
    listMeta.textContent = `No matches · ${total} on file`;
  } else {
    listMeta.textContent = `${rows.length} of ${total}`;
  }

  if (loading && !clients.length) {
    listEl.innerHTML = `
      <div class="p-6 text-sm text-ink/55" role="status">Loading members…</div>
    `;
    return;
  }

  if (lastError && !clients.length) {
    listEl.innerHTML = `
      <div class="p-6 space-y-3" role="alert">
        <p class="text-sm text-red-700">${escapeHtml(lastError)}</p>
        <button type="button" data-action="retry" class="min-h-11 px-4 rounded-full bg-navy text-white text-sm">Try again</button>
      </div>
    `;
    return;
  }

  if (!rows.length) {
    const empty = total
      ? "No members match that search."
      : "No memberships on file yet.";
    listEl.innerHTML = `<div class="p-6 text-sm text-ink/55">${empty}</div>`;
    return;
  }

  listEl.innerHTML = rows
    .map((client) => {
      const active = client.id === selectedId;
      const amount = formatCAD(client.totalAmount);
      return `
        <button
          type="button"
          role="listitem"
          data-id="${escapeHtml(client.id)}"
          class="w-full text-left px-4 sm:px-5 py-4 min-h-16 transition ${
            active ? "bg-navy-soft" : "bg-white hover:bg-cream/70"
          }"
        >
          <div class="flex items-start justify-between gap-3">
            <p class="font-medium text-navy truncate">${escapeHtml(client.fullName || "Unnamed")}</p>
            <span class="shrink-0 text-[11px] uppercase tracking-wide px-2 py-0.5 rounded-full border ${statusClass(
              client.leadStatus
            )}">${escapeHtml(client.leadStatus || "New")}</span>
          </div>
          <p class="mt-1 text-sm text-ink/65 truncate">${escapeHtml(client.phone || "—")} · ${escapeHtml(client.id)}</p>
          <p class="mt-1 text-sm text-ink/55 truncate">${escapeHtml(client.package || "IV Treatment Membership")}</p>
          <p class="mt-1 text-sm ${amount ? "text-navy font-medium" : "text-ink/40"}">${
            amount || "Total not set"
          }</p>
        </button>
      `;
    })
    .join("");
}

function fieldRow(label, value, extra = "") {
  return fieldHtml(label, value ? escapeHtml(value) : "", extra);
}

function fieldHtml(label, html, extra = "") {
  return `
    <div class="min-w-0">
      <dt class="text-xs uppercase tracking-wide text-ink/45 mb-1">${label}</dt>
      <dd class="text-sm text-ink break-words ${extra}">${html || "—"}</dd>
    </div>
  `;
}

function renderDetail() {
  const client = selectedClient();
  if (!client) {
    detailBody.innerHTML = `
      <div class="min-h-[16rem] flex flex-col items-center justify-center text-center px-4">
        <p class="text-navy text-xs font-medium tracking-[0.16em] uppercase mb-2">Client record</p>
        <h2 class="font-serif text-2xl text-navy mb-2">Select a member</h2>
        <p class="text-sm text-ink/55 max-w-sm">Search by name, phone, or submission ID, then tap a card to view details and set the total amount.</p>
      </div>
    `;
    return;
  }

  const notes = parseInternalNotes(client.internalNotes);
  const amountLabel = formatCAD(client.totalAmount) || "Not set";

  detailBody.innerHTML = `
    <div class="flex items-start justify-between gap-3 mb-5">
      <button
        type="button"
        data-action="back"
        class="lg:hidden min-h-11 px-3 -ml-2 rounded-full text-sm text-navy hover:bg-navy-soft"
      >
        ← Members
      </button>
      <div class="min-w-0 flex-1">
        <p class="text-navy text-xs font-medium tracking-[0.16em] uppercase mb-1">Member</p>
        <h2 class="font-serif text-2xl sm:text-3xl text-navy leading-tight break-words">${escapeHtml(
          client.fullName || "Unnamed"
        )}</h2>
        <p class="mt-1 text-sm text-ink/55">${escapeHtml(client.id)}</p>
      </div>
      <span class="shrink-0 text-xs uppercase tracking-wide px-2.5 py-1 rounded-full border ${statusClass(
        client.leadStatus
      )}">${escapeHtml(client.leadStatus || "New")}</span>
    </div>

    <dl class="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-6">
      ${fieldHtml(
        "Phone",
        client.phone
          ? `<a class="text-navy underline-offset-2 hover:underline" href="tel:${escapeHtml(
              digitsOnly(client.phone)
            )}">${escapeHtml(client.phone)}</a>`
          : ""
      )}
      ${fieldHtml(
        "Email",
        client.email
          ? `<a class="text-navy underline-offset-2 hover:underline break-all" href="mailto:${escapeHtml(
              client.email
            )}">${escapeHtml(client.email)}</a>`
          : ""
      )}
      ${fieldRow("Package", client.package || "IV Treatment Membership")}
      ${fieldRow("Term", client.term)}
      ${fieldRow("Lead status", client.leadStatus || "New")}
      ${fieldRow("Submitted", client.timestamp)}
      ${fieldRow("Date of birth", notes.dob)}
      ${fieldRow("Signature", notes.signature, "font-serif italic")}
      ${fieldRow("Heard from", notes.heardFrom || client.utmSource)}
    </dl>

    <section class="rounded-2xl border border-navy/10 bg-cream/70 p-4 sm:p-5 mb-5">
      <label for="total-amount" class="block text-sm font-medium text-navy mb-1">Total amount</label>
      <p class="text-xs text-ink/50 mb-3">Client value in CAD. Saved to the membership sheet.</p>
      <div class="flex flex-col sm:flex-row gap-3">
        <div class="relative flex-1 min-w-0">
          <span class="absolute left-3.5 top-1/2 -translate-y-1/2 text-ink/45">$</span>
          <input
            id="total-amount"
            type="text"
            inputmode="decimal"
            autocomplete="off"
            placeholder="0.00"
            value="${escapeHtml(client.totalAmount)}"
            class="w-full rounded-xl border border-mist bg-white pl-8 pr-14 py-3 text-base min-h-12 outline-none focus:border-navy focus:ring-2 focus:ring-navy/15"
          />
          <span class="absolute right-3.5 top-1/2 -translate-y-1/2 text-xs text-ink/45">CAD</span>
        </div>
        <button
          id="save-total"
          type="button"
          class="min-h-12 px-5 rounded-full bg-navy text-white text-sm font-medium hover:bg-navy-dark disabled:opacity-60 disabled:cursor-not-allowed"
        >
          Save total
        </button>
      </div>
      <p id="amount-status" class="mt-2 text-sm text-ink/55">Current: ${escapeHtml(amountLabel)}</p>
    </section>

    <section class="mb-5">
      <h3 class="font-serif text-lg text-navy mb-2">Health notes</h3>
      <p class="text-sm text-ink/80 whitespace-pre-wrap break-words rounded-2xl border border-mist bg-cream/50 p-4 min-h-16">${
        client.healthNotes ? escapeHtml(client.healthNotes) : "None on file."
      }</p>
    </section>

    <section>
      <h3 class="font-serif text-lg text-navy mb-2">Internal notes</h3>
      <p class="text-sm text-ink/80 whitespace-pre-wrap break-words rounded-2xl border border-mist bg-cream/50 p-4 min-h-16">${
        notes.extras.length
          ? escapeHtml(notes.extras.join(" · "))
          : "No additional notes."
      }</p>
    </section>
  `;
}

function render() {
  renderList();
  renderDetail();
  showList();
}

async function parseJsonResponse(response) {
  const text = await response.text();
  try {
    return JSON.parse(text);
  } catch (err) {
    return null;
  }
}

const MEMBERSHIP_VALIDATE_ERRORS = new Set([
  "Invalid name",
  "Invalid phone",
  "Invalid email",
  "Consent required",
]);
const LIST_UNAVAILABLE =
  "List is not available yet. Paste updated Code.gs, then Deploy → New version.";
const SAVE_UNAVAILABLE =
  "Save is not available yet. Paste updated Code.gs, then Deploy → New version.";
const REDEPLOY_HINT =
  "Desk API needs a New version deploy. Paste the latest Code.gs, then Deploy → Manage deployments → New version (same URL). No authorize/Allow step.";
// Google serves an HTML error page (not JSON) when the script itself fails to run.
const SCRIPT_ERROR_PAGE =
  "Apps Script returned an error page instead of data — paste the latest Code.gs into the editor (replacing all contents), then Deploy → Manage deployments → New version.";

function isAuthDenied(response, data) {
  if (response && response.status === 401) return true;
  const code = data && data.status != null ? Number(data.status) : 0;
  if (code === 401) return true;
  const error = data && data.error ? String(data.error) : "";
  return (
    /^sign in required$/i.test(error) ||
    /^unauthorized$/i.test(error) ||
    /^sign-in expired$/i.test(error)
  );
}

function needsRedeployHint(data) {
  const detail = data && data.detail ? String(data.detail) : "";
  const error = data && data.error ? String(data.error) : "";
  const blob = `${error} ${detail}`;
  return (
    /certs fetch failed/i.test(blob) ||
    /script\.external_request/i.test(blob) ||
    /UrlFetchApp/i.test(blob) ||
    /public keys unavailable/i.test(blob)
  );
}

function formatApiError(data) {
  const error = data && data.error ? String(data.error).trim() : "";
  const detail = data && data.detail ? String(data.detail).trim() : "";
  if (!error || MEMBERSHIP_VALIDATE_ERRORS.has(error)) return "";
  if (needsRedeployHint(data)) return REDEPLOY_HINT;
  return detail ? `${error} — ${detail}` : error;
}

function listErrorMessage(response, data) {
  const formatted = formatApiError(data);
  if (formatted) return formatted;
  if (isAuthDenied(response, data)) return "Sign in required";
  if (!data && response && response.status >= 400) return SCRIPT_ERROR_PAGE;
  if (response && response.status >= 400) {
    return `Could not load members (${response.status}).`;
  }
  return LIST_UNAVAILABLE;
}

function saveErrorMessage(response, data) {
  const formatted = formatApiError(data);
  if (formatted) return formatted;
  if (isAuthDenied(response, data)) return "Sign in required";
  if (!data && response && response.status >= 400) return SCRIPT_ERROR_PAGE;
  return SAVE_UNAVAILABLE;
}

function isClientList(data) {
  return Boolean(data && data.ok === true && Array.isArray(data.clients));
}

function isTotalSaved(data) {
  return Boolean(data && data.ok === true && !data.service && !Array.isArray(data.clients));
}

async function getIdToken(user, forceRefresh) {
  if (!user) return "";
  try {
    return await user.getIdToken(Boolean(forceRefresh));
  } catch (err) {
    if (!forceRefresh) {
      try {
        return await user.getIdToken(true);
      } catch (refreshErr) {
        return "";
      }
    }
    return "";
  }
}

function deskHeaders() {
  return { "Content-Type": "text/plain;charset=utf-8" };
}

async function deskPost(action, extra, idToken) {
  const payload = Object.assign({ action, idToken }, extra || {});
  const response = await fetch(`${WEBHOOK_URL}?action=${encodeURIComponent(action)}`, {
    method: "POST",
    redirect: "follow",
    headers: deskHeaders(),
    body: JSON.stringify(payload),
  });
  const data = await parseJsonResponse(response);
  return { response, data };
}

async function fetchClientList(idToken) {
  return deskPost("list", null, idToken);
}

async function loadMembers(options) {
  const opts = options || {};
  const user = opts.user || auth.currentUser;
  if (!user || loading) return;

  loading = true;
  lastError = "";
  refreshBtn.disabled = true;
  render();

  try {
    if (!WEBHOOK_URL) {
      throw new Error("Webhook URL is missing.");
    }

    let idToken = await getIdToken(user, Boolean(opts.forceRefresh));
    if (!idToken) {
      throw new Error("Sign in required");
    }

    let { response, data } = await fetchClientList(idToken);
    if (!isClientList(data) && isAuthDenied(response, data) && !opts.forceRefresh) {
      const fresh = await getIdToken(user, true);
      if (fresh) {
        ({ response, data } = await fetchClientList(fresh));
      }
    }

    if (!isClientList(data)) {
      throw new Error(listErrorMessage(response, data));
    }
    clients = data.clients;
    if (selectedId && !clients.some((client) => client.id === selectedId)) {
      selectedId = "";
    }
  } catch (err) {
    lastError =
      err && err.message
        ? err.message
        : "Could not load members. Check the desk URL and Apps Script deploy.";
    if (!clients.length) selectedId = "";
  } finally {
    loading = false;
    refreshBtn.disabled = false;
    render();
  }
}

async function saveTotal() {
  const client = selectedClient();
  const input = document.getElementById("total-amount");
  const status = document.getElementById("amount-status");
  const button = document.getElementById("save-total");
  if (!client || !input || saving) return;

  const parsed = parseAmount(input.value);
  if (parsed === null) {
    if (status) {
      status.textContent = "Enter a valid CAD amount, such as 50 or 1,280.00.";
      status.className = "mt-2 text-sm text-red-700";
    }
    input.focus();
    return;
  }

  saving = true;
  if (button) {
    button.disabled = true;
    button.textContent = "Saving…";
  }
  if (status) {
    status.textContent = "Saving total…";
    status.className = "mt-2 text-sm text-ink/55";
  }

  try {
    const user = auth.currentUser;
    if (!user) {
      throw new Error("Sign in required");
    }
    let idToken = await getIdToken(user, false);
    if (!idToken) {
      idToken = await getIdToken(user, true);
    }
    if (!idToken) {
      throw new Error("Sign in required");
    }
    let { response, data } = await deskPost(
      "updateTotal",
      { id: client.id, totalAmount: parsed },
      idToken
    );
    if (!isTotalSaved(data) && isAuthDenied(response, data)) {
      const fresh = await getIdToken(user, true);
      if (fresh) {
        ({ response, data } = await deskPost(
          "updateTotal",
          { id: client.id, totalAmount: parsed },
          fresh
        ));
      }
    }
    if (!isTotalSaved(data)) {
      throw new Error(saveErrorMessage(response, data));
    }
    client.totalAmount = parsed;
    input.value = "";
    renderList();
    if (status) {
      status.textContent = `Saved ${formatCAD(parsed) || "blank total"}.`;
      status.className = "mt-2 text-sm text-emerald-800";
    }
  } catch (err) {
    if (status) {
      status.textContent = err && err.message ? err.message : "Could not save the total.";
      status.className = "mt-2 text-sm text-red-700";
    }
  } finally {
    saving = false;
    if (button) {
      button.disabled = false;
      button.textContent = "Save total";
    }
  }
}

listEl.addEventListener("click", (event) => {
  const retry = event.target.closest("[data-action='retry']");
  if (retry) {
    loadMembers({ forceRefresh: true });
    return;
  }
  const button = event.target.closest("[data-id]");
  if (!button) return;
  selectedId = button.getAttribute("data-id") || "";
  render();
});

detailBody.addEventListener("click", (event) => {
  if (event.target.closest("[data-action='back']")) {
    selectedId = "";
    render();
    return;
  }
  if (event.target.closest("#save-total")) {
    saveTotal();
  }
});

detailBody.addEventListener("keydown", (event) => {
  if (event.key === "Enter" && event.target && event.target.id === "total-amount") {
    event.preventDefault();
    saveTotal();
  }
});

searchInput.addEventListener("input", () => {
  renderList();
});

refreshBtn.addEventListener("click", () => {
  loadMembers({ forceRefresh: true });
});

window.addEventListener("resize", () => {
  showList();
});

function goToLogin() {
  location.replace("login.html");
}

function showDesk(user) {
  signedIn = true;
  if (authLoading) authLoading.classList.add("hidden");
  if (deskApp) deskApp.classList.add("is-ready");
  if (signedInEmail) {
    signedInEmail.textContent = user.email || "";
    signedInEmail.classList.toggle("hidden", !user.email);
  }
}

if (signOutBtn) {
  signOutBtn.addEventListener("click", async () => {
    try {
      await auth.signOut();
      goToLogin();
    } catch (err) {
      lastError = "Could not sign out. Try again.";
      render();
    }
  });
}

auth.onAuthStateChanged((user) => {
  if (!user) {
    signedIn = false;
    goToLogin();
    return;
  }
  showDesk(user);
  loadMembers({ user });
});
